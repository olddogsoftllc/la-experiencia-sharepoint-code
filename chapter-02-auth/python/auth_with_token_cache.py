"""
auth_with_token_cache.py
Chapter 02: Authentication

Token Cache Authentication Example
Demonstrates persistent token caching for improved performance

Required environment variables:
    - TENANT_ID
    - CLIENT_ID
    - CLIENT_SECRET
"""

import os
import sys
import time
from typing import Dict, Optional, Any
from datetime import datetime, timedelta
import requests


class TokenCache:
    """Simple in-memory token cache implementation"""

    def __init__(self, token_lifetime_minutes: int = 55):
        self._cache: Dict[str, Any] = {}
        self.token_lifetime_minutes = token_lifetime_minutes

    def get(self, key: str) -> Optional[str]:
        """Gets a token from cache if valid"""
        cached = self._cache.get(key)
        if not cached:
            return None

        if datetime.now() < cached['expiry']:
            print("Using cached access token")
            return cached['token']

        # Token expired, remove from cache
        del self._cache[key]
        return None

    def set(self, key: str, token: str) -> None:
        """Stores a token in cache"""
        self._cache[key] = {
            'token': token,
            'expiry': datetime.now() + timedelta(minutes=self.token_lifetime_minutes)
        }
        print("New access token acquired and cached")

    def clear(self) -> None:
        """Clears the token cache"""
        self._cache.clear()
        print("Token cache cleared")

    def get_status(self, key: str) -> Dict[str, Any]:
        """Gets cache status information"""
        cached = self._cache.get(key)
        return {
            'has_token': key in self._cache,
            'is_valid': cached is not None and datetime.now() < cached['expiry'],
            'expiry_time': cached['expiry'].isoformat() if cached else None
        }


class AuthWithTokenCache:
    """Authentication with token caching support"""

    def __init__(self):
        self.tenant_id = os.environ.get('TENANT_ID')
        self.client_id = os.environ.get('CLIENT_ID')
        self.client_secret = os.environ.get('CLIENT_SECRET')
        self.cache = TokenCache()
        self.cache_key = 'graph_access_token'

        self._validate_config()

    def _validate_config(self) -> None:
        """Validates configuration"""
        required = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET']
        missing = [key for key in required if not os.environ.get(key)]

        if missing:
            raise ValueError(f"Missing required environment variables: {', '.join(missing)}")

    def get_access_token(self) -> str:
        """
        Gets token from cache or requests new one

        Returns:
            str: Access token
        """
        # Check cache first
        cached_token = self.cache.get(self.cache_key)
        if cached_token:
            return cached_token

        # Request new token
        try:
            token_endpoint = (
                f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
            )
            request_data = {
                'client_id': self.client_id,
                'client_secret': self.client_secret,
                'scope': 'https://graph.microsoft.com/.default',
                'grant_type': 'client_credentials'
            }

            response = requests.post(
                token_endpoint,
                data=request_data,
                headers={'Content-Type': 'application/x-www-form-urlencoded'}
            )
            response.raise_for_status()

            token = response.json()['access_token']
            self.cache.set(self.cache_key, token)

            return token

        except requests.exceptions.RequestException as e:
            print(f"Failed to acquire token: {e}")
            raise

    def get_authenticated_headers(self) -> Dict[str, str]:
        """Gets authenticated headers with token caching"""
        token = self.get_access_token()
        return {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }

    def clear_cache(self) -> None:
        """Clears the token cache"""
        self.cache.clear()

    def get_cache_status(self) -> Dict[str, Any]:
        """Gets cache status"""
        return self.cache.get_status(self.cache_key)


def main():
    """Main execution function"""
    try:
        print("=== Token Cache Authentication Example ===\n")

        auth = AuthWithTokenCache()

        # First call - will acquire new token
        print("First call (new token):")
        token1 = auth.get_access_token()
        print(f"Token acquired: {token1[:20]}...")

        # Check cache status
        status = auth.get_cache_status()
        print(f"Cache status: HasToken={status['has_token']}, IsValid={status['is_valid']}\n")

        # Second call - should use cached token
        print("Second call (from cache):")
        token2 = auth.get_access_token()
        print(f"Token from cache: {token2[:20]}...")

        # Verify tokens match
        print(f"\nTokens match: {token1 == token2}")

        # Clear cache
        auth.clear_cache()
        print(f"After clear - HasToken: {auth.get_cache_status()['has_token']}")

        print("\nToken caching example completed!")

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
