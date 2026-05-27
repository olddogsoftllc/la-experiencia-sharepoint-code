"""
document_operations.py
Chapter 04: Documents

SharePoint Document Operations Example
Demonstrates upload, download, and search operations for documents

Required environment variables:
    - TENANT_ID
    - CLIENT_ID
    - CLIENT_SECRET
"""

import os
import sys
from typing import Dict, List, Optional
import requests


class DocumentOperations:
    """SharePoint Document Operations handler"""

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

    def _get_headers(self, content_type: str = 'application/json') -> Dict[str, str]:
        """Gets authenticated headers"""
        token = self._get_access_token()
        return {
            'Authorization': f'Bearer {token}',
            'Content-Type': content_type
        }

    def upload_file(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
        destination_file_name: str
    ) -> Dict:
        """
        Uploads a file to SharePoint

        Args:
            site_id: The site ID
            drive_id: The drive ID
            file_path: Local path to the file
            destination_file_name: Name for the file in SharePoint

        Returns:
            Uploaded file metadata
        """
        try:
            print(f"Uploading file: {file_path}")

            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File not found: {file_path}")

            with open(file_path, 'rb') as f:
                file_content = f.read()

            encoded_name = requests.utils.quote(destination_file_name)
            headers = self._get_headers('application/octet-stream')
            url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/"
                f"{drive_id}/items/root:/{encoded_name}:/content"
            )

            response = requests.put(url, headers=headers, data=file_content)
            response.raise_for_status()

            result = response.json()
            print("File uploaded successfully:")
            print(f"  ID: {result['id']}")
            print(f"  Name: {result['name']}")
            print(f"  Size: {result['size']} bytes")
            print(f"  Web URL: {result['webUrl']}")

            return result

        except requests.exceptions.RequestException as e:
            print(f"Error uploading file: {e}")
            raise

    def download_file(
        self,
        site_id: str,
        drive_id: str,
        item_id: str,
        download_path: str
    ) -> None:
        """
        Downloads a file from SharePoint

        Args:
            site_id: The site ID
            drive_id: The drive ID
            item_id: The item ID of the file
            download_path: Local path to save the file
        """
        try:
            print(f"Downloading file to: {download_path}")

            headers = self._get_headers()
            url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/"
                f"{drive_id}/items/{item_id}/content"
            )

            response = requests.get(url, headers=headers)
            response.raise_for_status()

            with open(download_path, 'wb') as f:
                f.write(response.content)

            print(f"File downloaded successfully to: {download_path}")

        except requests.exceptions.RequestException as e:
            print(f"Error downloading file: {e}")
            raise

    def search_files(self, query: str) -> Dict:
        """
        Searches for files across SharePoint

        Args:
            query: Search query

        Returns:
            Search results
        """
        try:
            print(f"Searching for files: '{query}'")

            headers = self._get_headers()
            search_body = {
                "requests": [{
                    "entityTypes": ["driveItem"],
                    "query": {
                        "queryString": query
                    }
                }]
            }

            response = requests.post(
                'https://graph.microsoft.com/v1.0/search/query',
                headers=headers,
                json=search_body
            )
            response.raise_for_status()

            results = response.json()
            print("Search results:")
            print("-" * 80)

            for result in results.get('value', []):
                for container in result.get('hitsContainers', []):
                    for hit in container.get('hits', []):
                        resource = hit.get('resource', {})
                        print(f"Name: {resource.get('name')}")
                        print(f"  Web URL: {resource.get('webUrl')}")
                        print(f"  Size: {resource.get('size')} bytes")
                        print("-" * 80)

            return results

        except requests.exceptions.RequestException as e:
            print(f"Error searching files: {e}")
            raise

    def list_files_in_folder(
        self,
        site_id: str,
        drive_id: str,
        folder_path: str = ""
    ) -> List[Dict]:
        """
        Lists files in a specific folder

        Args:
            site_id: The site ID
            drive_id: The drive ID
            folder_path: Path to the folder

        Returns:
            List of items
        """
        try:
            print(f"Listing files in folder: {folder_path or 'root'}")

            headers = self._get_headers()

            url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/"
                f"{drive_id}/root/children"
            )

            if folder_path:
                encoded_path = requests.utils.quote(folder_path)
                url = (
                    f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/"
                    f"{drive_id}/items/root:/{encoded_path}:/children"
                )

            response = requests.get(url, headers=headers)
            response.raise_for_status()

            items = response.json().get('value', [])
            print(f"Found {len(items)} items:")
            print("-" * 80)

            for item in items:
                item_type = "Folder" if item.get('folder') else "File"
                print(f"{item_type}: {item['name']}")
                print(f"  ID: {item['id']}")
                print(f"  Size: {item.get('size', 0)} bytes")
                print(f"  Last Modified: {item.get('lastModifiedDateTime')}")
                print("-" * 80)

            return items

        except requests.exceptions.RequestException as e:
            print(f"Error listing files: {e}")
            raise

    def get_file_metadata(
        self,
        site_id: str,
        drive_id: str,
        item_id: str
    ) -> Dict:
        """
        Gets file metadata

        Args:
            site_id: The site ID
            drive_id: The drive ID
            item_id: The item ID

        Returns:
            File metadata
        """
        try:
            print(f"Fetching metadata for item: {item_id}")

            headers = self._get_headers()
            url = (
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/"
                f"{drive_id}/items/{item_id}"
            )

            response = requests.get(url, headers=headers)
            response.raise_for_status()

            item = response.json()
            print("File metadata:")
            print(f"  Name: {item['name']}")
            print(f"  ID: {item['id']}")
            print(f"  Size: {item.get('size', 0)} bytes")
            print(f"  Created: {item.get('createdDateTime')}")
            print(f"  Modified: {item.get('lastModifiedDateTime')}")
            print(f"  Web URL: {item.get('webUrl')}")

            return item

        except requests.exceptions.RequestException as e:
            print(f"Error getting file metadata: {e}")
            raise


def main():
    """Main execution function"""
    try:
        print("=== SharePoint Document Operations Example ===\n")

        doc_ops = DocumentOperations()

        print("Document operations class initialized successfully!")
        print("Use the methods to perform upload, download, and search operations.")

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
