"""
graph_auth_example.py
Chapter 02: Authentication

Microsoft Graph Client Credentials Authentication Example
Demonstrates authenticating to Microsoft Graph using client credentials flow

Required environment variables:
    - TENANT_ID
    - CLIENT_ID
    - CLIENT_SECRET
"""

import os
import sys
from typing import Dict, Optional
import requests
from urllib.parse import urlencode


class GraphAuthConfig:
    """Configuration for Microsoft Graph authentication"""

    def __init__(self):
        self.tenant_id = os.environ.get('TENANT_ID')
        self.client_id = os.environ.get('CLIENT_ID')
        self.client_secret = os.environ.get('CLIENT_SECRET')

        self._validate_config()

    def _validate_config(self) -> None:
        """Validates that all required environment variables are present"""
        required = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET']
        missing = [key for key in required if not os.environ.get(key)]

        if missing:
            raise ValueError(f"Missing required environment variables: {', '.join(missing)}")


class GraphAuthExample:
    """Microsoft Graph Authentication handler using client credentials"""

    def __init__(self):
        self.config = GraphAuthConfig()
        self.token_endpoint = (
            f"https://login.microsoftonline.com/{self.config.tenant_id}/oauth2/v2.0/token"
        )
        self.scope = "https://graph.microsoft.com/.default"

    def get_access_token(self) -> str:
        """
        Acquires an access token using client credentials flow

        Returns:
            str: Access token for Microsoft Graph API
        """
        try:
            request_data = {
                'client_id': self.config.client_id,
                'client_secret': self.config.client_secret,
                'scope': self.scope,
                'grant_type': 'client_credentials'
            }

            response = requests.post(
                self.token_endpoint,
                data=request_data,
                headers={'Content-Type': 'application/x-www-form-urlencoded'}
            )
            response.raise_for_status()

            print("Successfully authenticated to Microsoft Graph")
            return response.json()['access_token']

        except requests.exceptions.RequestException as e:
            print(f"Authentication failed: {e}")
            raise Exception("Failed to acquire access token") from e

    def get_authenticated_headers(self) -> Dict[str, str]:
        """
        Creates authenticated API client headers

        Returns:
            Dict[str, str]: Headers dictionary with authorization
        """
        token = self.get_access_token()
        return {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }

    def test_connection(self) -> Optional[Dict]:
        """
        Tests the connection by retrieving organization details

        Returns:
            Optional[Dict]: Organization information or None if failed
        """
        try:
            headers = self.get_authenticated_headers()
            response = requests.get(
                'https://graph.microsoft.com/v1.0/organization',
                headers=headers
            )
            response.raise_for_status()

            org = response.json()['value'][0]
            print(f"Connected to tenant: {org['displayName']}")
            return response.json()

        except requests.exceptions.RequestException as e:
            print(f"Connection test failed: {e}")
            raise


def main():
    """Main execution function"""
    try:
        print("=== Microsoft Graph Authentication Example ===")

        auth = GraphAuthExample()
        auth.test_connection()

        print("\nAuthentication completed successfully!")

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
