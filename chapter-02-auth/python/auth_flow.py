"""
Capítulo 2: Autenticación
Ejemplo de Client Credentials con MSAL
"""

import os
from msgraph.core import GraphClient
from azure.identity import ClientSecretCredential


def main():
    print("=== Capítulo 2: Autenticación ===\n")

    tenant_id = os.getenv('TENANT_ID')
    client_id = os.getenv('CLIENT_ID')
    client_secret = os.getenv('CLIENT_SECRET')

    if not all([tenant_id, client_id, client_secret]):
        print("Error: Faltan variables de entorno")
        return

    try:
        credential = ClientSecretCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret
        )

        client = GraphClient(credential=credential)
        org = client.get('/organization').json()

        print("✓ Conexión exitosa")
        print(f"  Tenant: {org['value'][0]['displayName']}")
        print(f"  ID: {org['value'][0]['id']}")

    except Exception as e:
        print(f"✗ Error: {e}")


if __name__ == '__main__':
    main()
