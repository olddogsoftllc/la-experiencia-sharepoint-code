"""
Chapter 2: Authentication
Client Credentials example with MSAL
"""

import os
from msgraph.core import GraphClient
from azure.identity import ClientSecretCredential


def main():
    print("=== Chapter 2: Authentication ===\n")

    tenant_id = os.getenv('TENANT_ID')
    client_id = os.getenv('CLIENT_ID')
    client_secret = os.getenv('CLIENT_SECRET')

    if not all([tenant_id, client_id, client_secret]):
        print("Error: Missing environment variables")
        return

    try:
        credential = ClientSecretCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret
        )

        client = GraphClient(credential=credential)
        org = client.get('/organization').json()

        print("✓ Connection successful")
        print(f"  Tenant: {org['value'][0]['displayName']}")
        print(f"  ID: {org['value'][0]['id']}")

    except Exception as e:
        print(f"✗ Error: {e}")


if __name__ == '__main__':
    main()
