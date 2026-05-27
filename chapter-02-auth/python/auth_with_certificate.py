"""
auth_with_certificate.py
Chapter 02: Authentication

Certificate-Based Authentication Example
Demonstrates secure authentication using X.509 certificates

Required environment variables:
    - TENANT_ID
    - CLIENT_ID
    - CERTIFICATE_PATH (or CERTIFICATE_THUMBPRINT with certificate store access)
    - CERTIFICATE_PASSWORD (optional, for encrypted certificates)
"""

import os
import sys
import base64
import hashlib
from typing import Dict, Optional
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.serialization import pkcs12
from cryptography.hazmat.backends import default_backend
import requests
from msal import ConfidentialClientApplication


class CertificateLoader:
    """Certificate loading utility"""

    @staticmethod
    def load_from_file(file_path: str, password: Optional[str] = None) -> Dict:
        """
        Loads certificate from file

        Args:
            file_path: Path to certificate file (.pfx or .pem)
            password: Optional password for encrypted certificate

        Returns:
            Dict with private_key and thumbprint
        """
        try:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"Certificate file not found: {file_path}")

            with open(file_path, 'rb') as f:
                cert_data = f.read()

            # Decode password if provided
            cert_password = password.encode() if password else None

            # Load PKCS12 certificate
            private_key, certificate, _ = pkcs12.load_key_and_certificates(
                cert_data, cert_password, default_backend()
            )

            # Calculate thumbprint (SHA1 of DER-encoded certificate)
            cert_der = certificate.public_bytes(serialization.Encoding.DER)
            thumbprint = hashlib.sha1(cert_der).hexdigest().upper()

            # Export private key in PEM format
            private_key_pem = private_key.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.TraditionalOpenSSL,
                encryption_algorithm=serialization.NoEncryption()
            ).decode('utf-8')

            print("Certificate loaded from file")

            return {
                'private_key': private_key_pem,
                'thumbprint': thumbprint,
                'certificate': certificate
            }

        except Exception as e:
            print(f"Failed to load certificate from file: {e}")
            raise

    @staticmethod
    def validate_certificate(cert_data: Dict) -> None:
        """
        Validates certificate details

        Args:
            cert_data: Certificate data dictionary
        """
        cert = cert_data['certificate']

        print("\nCertificate Details:")
        print(f"  Subject: {cert.subject}")
        print(f"  Issuer: {cert.issuer}")
        print(f"  Thumbprint: {cert_data['thumbprint']}")
        print(f"  Valid From: {cert.not_valid_before}")
        print(f"  Valid Until: {cert.not_valid_after}")
        print(f"  Has Private Key: {cert_data['private_key'] is not None}")

        # Check expiration
        from datetime import datetime
        now = datetime.now()

        if cert.not_valid_after < now:
            raise ValueError("Certificate has expired")

        if cert.not_valid_before > now:
            raise ValueError("Certificate is not yet valid")

        if not cert_data['private_key']:
            raise ValueError("Certificate does not have a private key")

        print("\nCertificate validation passed")


class AuthWithCertificate:
    """Certificate-based authentication handler"""

    def __init__(self):
        self.tenant_id = os.environ.get('TENANT_ID')
        self.client_id = os.environ.get('CLIENT_ID')
        self.certificate_path = os.environ.get('CERTIFICATE_PATH')
        self.certificate_password = os.environ.get('CERTIFICATE_PASSWORD')

        self._validate_config()

    def _validate_config(self) -> None:
        """Validates configuration"""
        if not self.tenant_id or not self.client_id:
            raise ValueError("TENANT_ID and CLIENT_ID environment variables are required")

        if not self.certificate_path:
            raise ValueError("CERTIFICATE_PATH environment variable is required")

    def get_access_token(self) -> str:
        """
        Acquires access token using certificate authentication

        Returns:
            str: Access token
        """
        try:
            # Load certificate
            cert_data = CertificateLoader.load_from_file(
                self.certificate_path,
                self.certificate_password
            )
            CertificateLoader.validate_certificate(cert_data)

            # Create MSAL application
            authority = f"https://login.microsoftonline.com/{self.tenant_id}"

            app = ConfidentialClientApplication(
                client_id=self.client_id,
                authority=authority,
                client_credential={
                    'private_key': cert_data['private_key'],
                    'thumbprint': cert_data['thumbprint']
                }
            )

            # Acquire token
            result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

            if 'access_token' not in result:
                error = result.get('error_description', 'Unknown error')
                raise Exception(f"Failed to acquire token: {error}")

            print("Successfully authenticated using certificate")
            return result['access_token']

        except Exception as e:
            print(f"Certificate authentication failed: {e}")
            raise

    def get_authenticated_headers(self) -> Dict[str, str]:
        """Gets authenticated headers"""
        token = self.get_access_token()
        return {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }


def main():
    """Main execution function"""
    try:
        print("=== Certificate-Based Authentication Example ===\n")

        auth = AuthWithCertificate()
        token = auth.get_access_token()

        print(f"\nToken acquired: {token[:30]}...")
        print("\nCertificate authentication completed successfully!")

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
