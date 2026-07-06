"""
document_operations.py
Chapter 04: Documents

SharePoint Document Operations Example
Demonstrates upload, download, and search operations for documents.

Usa el módulo de auth compartido (common/laexperiencia_sharepoint): el access token
se inyecta por constructor (DI), no se obtiene dentro de la clase.
"""

import os
import sys
from typing import Dict, List, Optional
import requests

from laexperiencia_sharepoint import get_access_token

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class DocumentOperations:
    """SharePoint Document Operations handler."""

    def __init__(self, access_token: str):
        if not access_token:
            raise ValueError("Se requiere un access token para DocumentOperations.")
        self.access_token = access_token

    def _get_headers(self, content_type: str = 'application/json') -> Dict[str, str]:
        return {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': content_type,
        }

    def list_drives(self, site_id: str) -> List[Dict]:
        """Lists the document libraries (drives) of a site. Solo lectura."""
        try:
            print(f"Listing document libraries of site: {site_id}")
            url = f"{GRAPH_BASE}/sites/{site_id}/drives"
            response = requests.get(url, headers=self._get_headers())
            response.raise_for_status()
            drives = response.json().get('value', [])
            print(f"Found {len(drives)} libraries:")
            print("-" * 80)
            for drive in drives:
                print(f"Library: {drive.get('name')}")
                print(f"  ID: {drive.get('id')}")
                print(f"  Type: {drive.get('driveType')}")
                print("-" * 80)
            return drives
        except requests.exceptions.RequestException as e:
            print(f"Error listing libraries: {e}")
            raise

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
    """Construye el token con el módulo común y lista las bibliotecas del sitio book-test."""
    try:
        print("=== SharePoint Document Operations Example ===\n")
        token = get_access_token()
        doc_ops = DocumentOperations(token)

        hostname = os.environ.get("SHAREPOINT_HOSTNAME", "olddogsoft1.sharepoint.com")
        site_path = os.environ.get("SHAREPOINT_SITE_PATH", "book-test")

        # Resolver el sitio por path para obtener su ID (el path form no sirve para /drives).
        site_url = (
            f"{GRAPH_BASE}/sites/"
            f"{requests.utils.quote(hostname, safe='')}:/sites/"
            f"{requests.utils.quote(site_path, safe='')}"
        )
        site = requests.get(site_url, headers=doc_ops._get_headers())
        site.raise_for_status()
        site_id = site.json()["id"]

        doc_ops.list_drives(site_id)

        print("\nDocument operations completed successfully!")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
