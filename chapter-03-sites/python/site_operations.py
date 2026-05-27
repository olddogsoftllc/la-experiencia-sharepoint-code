"""
site_operations.py
Chapter 03: Sites

SharePoint Site Operations Example
Demonstrates listing, creating, and retrieving SharePoint sites

Required environment variables:
    - TENANT_ID
    - CLIENT_ID
    - CLIENT_SECRET
"""

import os
import sys
from typing import Dict, List, Optional
import requests


class SiteOperations:
    """SharePoint Site Operations handler"""

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

    def list_sites(self) -> List[Dict]:
        """
        Lists all sites in the organization

        Returns:
            List of site dictionaries
        """
        try:
            print("Fetching all sites...\n")

            headers = self._get_headers()
            response = requests.get(
                'https://graph.microsoft.com/v1.0/sites',
                headers=headers
            )
            response.raise_for_status()

            sites = response.json()['value']
            print(f"Found {len(sites)} sites:")
            print("-" * 80)

            for site in sites:
                print(f"Name: {site['name']}")
                print(f"  ID: {site['id']}")
                print(f"  Web URL: {site['webUrl']}")
                print(f"  Display Name: {site.get('displayName', 'N/A')}")
                print("-" * 80)

            return sites

        except requests.exceptions.RequestException as e:
            print(f"Error listing sites: {e}")
            raise

    def get_site(self, hostname: str, site_path: str) -> Optional[Dict]:
        """
        Gets a specific site by hostname and site path

        Args:
            hostname: The hostname of the site (e.g., contoso.sharepoint.com)
            site_path: The site path (e.g., sites/marketing)

        Returns:
            Site dictionary or None
        """
        try:
            print(f"Fetching site: {hostname}/sites/{site_path}")

            headers = self._get_headers()
            encoded_path = requests.utils.quote(f"sites/{site_path}")
            url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{encoded_path}"

            response = requests.get(url, headers=headers)
            response.raise_for_status()

            site = response.json()
            print("\nSite found:")
            print(f"  Name: {site['name']}")
            print(f"  ID: {site['id']}")
            print(f"  Web URL: {site['webUrl']}")
            print(f"  Description: {site.get('description', 'N/A')}")

            return site

        except requests.exceptions.RequestException as e:
            print(f"Error getting site: {e}")
            raise

    def get_site_by_id(self, site_id: str) -> Optional[Dict]:
        """
        Gets site by its unique identifier

        Args:
            site_id: The site ID

        Returns:
            Site dictionary or None
        """
        try:
            print(f"Fetching site by ID: {site_id}")

            headers = self._get_headers()
            response = requests.get(
                f"https://graph.microsoft.com/v1.0/sites/{site_id}",
                headers=headers
            )
            response.raise_for_status()

            site = response.json()
            print("\nSite found:")
            print(f"  Name: {site['name']}")
            print(f"  Web URL: {site['webUrl']}")

            return site

        except requests.exceptions.RequestException as e:
            print(f"Error getting site by ID: {e}")
            raise

    def get_root_site(self) -> Optional[Dict]:
        """
        Gets the root site of the organization

        Returns:
            Site dictionary or None
        """
        try:
            print("Fetching root site...")

            headers = self._get_headers()
            response = requests.get(
                'https://graph.microsoft.com/v1.0/sites/root',
                headers=headers
            )
            response.raise_for_status()

            site = response.json()
            print("\nRoot site:")
            print(f"  Name: {site['name']}")
            print(f"  ID: {site['id']}")
            print(f"  Web URL: {site['webUrl']}")

            return site

        except requests.exceptions.RequestException as e:
            print(f"Error getting root site: {e}")
            raise

    def search_sites(self, keyword: str) -> List[Dict]:
        """
        Searches for sites by keyword

        Args:
            keyword: Search keyword

        Returns:
            List of matching site dictionaries
        """
        try:
            print(f"Searching for sites with keyword: '{keyword}'")

            headers = self._get_headers()
            response = requests.get(
                f"https://graph.microsoft.com/v1.0/sites?search={keyword}",
                headers=headers
            )
            response.raise_for_status()

            sites = response.json()['value']
            print(f"\nFound {len(sites)} matching sites:")
            print("-" * 80)

            for site in sites:
                print(f"Name: {site['name']}")
                print(f"  Web URL: {site['webUrl']}")
                print(f"  Description: {site.get('description', 'N/A')}")
                print("-" * 80)

            return sites

        except requests.exceptions.RequestException as e:
            print(f"Error searching sites: {e}")
            raise


def main():
    """Main execution function"""
    try:
        print("=== SharePoint Site Operations Example ===\n")

        site_ops = SiteOperations()

        # Get root site
        site_ops.get_root_site()

        # List all sites
        site_ops.list_sites()

        print("\nSite operations completed successfully!")

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
