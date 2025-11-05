import os
from dataclasses import dataclass
from io import BytesIO
from urllib.parse import quote

import msal
import requests
from dotenv import load_dotenv

load_dotenv(dotenv_path='keys.env')


@dataclass
class ItemInfo:
    """Data class for file or folder information from SharePoint"""
    name: str
    path: str
    size: int
    modified: str
    id: str
    webUrl: str


class SharePointManager:
    # Base URL for Microsoft Graph API
    GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"

    def __init__(self, tenant_id, client_id, client_secret, site_name):
        """
        Initialize SharePoint connection using MSAL

        Args:
            tenant_id: Azure AD Tenant ID
            client_id: Azure AD Application (client) ID
            client_secret: Azure AD Application client secret
            site_name: SharePoint site name (e.g., 'yourtenant.sharepoint.com' or just 'yourtenant')
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret

        # Ensure site_name has the correct format
        if not site_name.endswith('.sharepoint.com'):
            self.site_name = f"{site_name}.sharepoint.com"
        else:
            self.site_name = site_name

        self.access_token = None
        self.site_id = None
        self.drive_id = None

        # Authenticate and get access token
        self._authenticate()

    def _authenticate(self):
        """
        Authenticate using MSAL and get access token
        """
        authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        scope = ["https://graph.microsoft.com/.default"]

        # Create a confidential client application
        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=authority,
            client_credential=self.client_secret
        )

        # Acquire token
        result = app.acquire_token_for_client(scopes=scope)

        if "access_token" in result:
            self.access_token = result["access_token"]
            print("✓ Authentication successful")
        else:
            error_msg = result.get("error_description", result.get("error"))
            raise Exception(f"Authentication failed: {error_msg}")

    def _get_headers(self):
        """Get request headers with access token"""
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

    def _get_site_url(self, site_path=""):
        """Build site URL"""
        if site_path:
            return f"{self.GRAPH_API_BASE}/sites/{self.site_name}:{site_path}"
        else:
            return f"{self.GRAPH_API_BASE}/sites/{self.site_name}"

    def _get_drives_url(self):
        """Build drives list URL"""
        return f"{self.GRAPH_API_BASE}/sites/{self.site_id}/drives"

    def _get_drive_root_url(self):
        """Build drive root URL"""
        return f"{self.GRAPH_API_BASE}/sites/{self.site_id}/drives/{self.drive_id}/root"

    def _get_drive_item_url(self, encoded_path):
        """Build drive item URL with path"""
        return f"{self._get_drive_root_url()}:/{encoded_path}"

    def _get_drive_item_content_url(self, encoded_path):
        """Build drive item content URL for upload"""
        return f"{self._get_drive_item_url(encoded_path)}:/content"

    def _get_drive_children_url(self, folder_path=""):
        """Build drive children URL for listing items in a folder"""
        if folder_path:
            encoded_path = quote(folder_path)
            return f"{self._get_drive_item_url(encoded_path)}:/children"
        else:
            return f"{self._get_drive_root_url()}/children"

    def _get_item_id_by_path(self, item_path):
        """
        Get the SharePoint item ID for a file or folder by its path

        Args:
            item_path: Path to the item (e.g., 'folder/file.txt' or 'folder_name')

        Returns:
            Item ID string
        """
        try:
            encoded_path = quote(item_path)
            url = self._get_drive_item_url(encoded_path)
            response = requests.get(url, headers=self._get_headers())
            response.raise_for_status()

            item_id = response.json().get("id")
            if not item_id:
                raise Exception(f"Could not retrieve ID for item: {item_path}")

            return item_id

        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                raise Exception(f"Item not found: {item_path}")
            else:
                raise Exception(f"Error getting item ID: {e.response.status_code} - {e.response.text}")
        except Exception as e:
            raise Exception(f"Error retrieving item ID: {str(e)}")

    def get_site_id(self, site_path=""):
        """
        Get SharePoint site ID

        Args:
            site_path: Site path (e.g., '/sites/yoursite' or empty for root site)

        Returns:
            Site ID
        """
        url = self._get_site_url(site_path)
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()

        self.site_id = response.json()["id"]
        print(f"✓ Site ID retrieved: {self.site_id}")
        return self.site_id

    def get_drive_id(self, drive_name="Documenten"):
        """
        Get document library (drive) ID

        Args:
            drive_name: Name of the document library (default: 'Documents')

        Returns:
            Drive ID
        """
        if not self.site_id:
            raise Exception("Site ID not set. Call get_site_id() first.")

        url = self._get_drives_url()
        response = requests.get(url, headers=self._get_headers())
        response.raise_for_status()

        drives = response.json()["value"]

        # Find the drive by name
        for drive in drives:
            if drive["name"] == drive_name:
                self.drive_id = drive["id"]
                print(f"✓ Drive ID retrieved: {self.drive_id}")
                return self.drive_id

        # If not found, use the default drive
        if drives:
            self.drive_id = drives[0]["id"]
            print(f"✓ Using default drive ID: {self.drive_id}")
            return self.drive_id

        raise Exception(f"No drives found or drive '{drive_name}' not found")

    def download_item(self, item_path, local_path=None):
        """
        Download a file or folder from SharePoint (auto-detects type)

        Args:
            item_path: Path to the item in SharePoint (e.g., 'folder/data.csv' or 'folder_name')
            local_path: Local path to save the item (optional)
                        - For files: full file path or None for in-memory
                        - For folders: directory path or None for current directory

        Returns:
            For files: BytesIO object if local_path is None, otherwise the local file path
            For folders: Path to the downloaded folder
        """
        if not self.drive_id:
            raise Exception("Drive ID not set. Call get_drive_id() first.")

        try:
            # Encode the item path
            encoded_path = quote(item_path)

            # Get item metadata to determine if it's a file or folder
            url = self._get_drive_item_url(encoded_path)
            response = requests.get(url, headers=self._get_headers())
            response.raise_for_status()

            item_info = response.json()

            # Check if it's a file or folder
            if "file" in item_info:
                # It's a file - use download_file logic
                return self._download_file_internal(item_path, local_path, item_info)
            elif "folder" in item_info:
                # It's a folder - use download_folder logic
                return self._download_folder_internal(item_path, local_path)
            else:
                raise Exception(f"Unknown item type for: {item_path}")

        except requests.exceptions.HTTPError as e:
            print(f"✗ Error downloading item: {e.response.status_code} - {e.response.text}")
            raise
        except Exception as e:
            print(f"✗ Error downloading item: {str(e)}")
            raise

    def _download_file_internal(self, file_path, local_path, item_info):
        """Internal method to download a file (used by download_item)"""
        download_url = item_info["@microsoft.graph.downloadUrl"]

        # Download file content
        file_response = requests.get(download_url)
        file_response.raise_for_status()

        if local_path:
            # Save to local file
            with open(local_path, 'wb') as local_file:
                local_file.write(file_response.content)
            print(f"✓ File downloaded successfully to: {local_path}")
            return local_path
        else:
            # Return as BytesIO object for in-memory processing
            print("✓ File downloaded successfully to memory")
            return BytesIO(file_response.content)

    def _download_folder_internal(self, folder_path, local_directory):
        """Internal method to download a folder (used by download_item)"""
        # Get folder name from path
        folder_name = folder_path.split("/")[-1]

        # Set default local directory
        if local_directory is None:
            local_directory = folder_name

        # Create local directory if it doesn't exist
        os.makedirs(local_directory, exist_ok=True)

        print(f"✓ Starting download of folder: {folder_path}")

        def download_folder_recursive(sharepoint_path, local_path):
            """Recursively download folder and its contents"""
            # Build the URL for current folder
            url = self._get_drive_children_url(sharepoint_path)

            # Get all items in the current folder
            response = requests.get(url, headers=self._get_headers())
            response.raise_for_status()

            items = response.json().get("value", [])

            for item in items:
                item_name = item.get("name", "")

                if "file" in item:
                    # Download file
                    file_download_url = item.get("@microsoft.graph.downloadUrl")
                    if file_download_url:
                        file_response = requests.get(file_download_url)
                        file_response.raise_for_status()

                        # Save file to local path
                        local_file_path = os.path.join(local_path, item_name)
                        with open(local_file_path, 'wb') as local_file:
                            local_file.write(file_response.content)
                        print(f"  ✓ Downloaded file: {item_name}")

                elif "folder" in item:
                    # Create local subfolder
                    local_subfolder = os.path.join(local_path, item_name)
                    os.makedirs(local_subfolder, exist_ok=True)

                    # Recursively download subfolder
                    new_sharepoint_path = f"{sharepoint_path}/{item_name}" if sharepoint_path else item_name
                    download_folder_recursive(new_sharepoint_path, local_subfolder)

        # Start the recursive download
        download_folder_recursive(folder_path, local_directory)

        print(f"✓ Folder downloaded successfully to: {local_directory}")
        return local_directory

    def download_file(self, file_path, local_path=None):
        """
        Download a file from SharePoint

        Deprecated: Use download_item() instead. This method is kept for backward compatibility.

        Args:
            file_path: Path to file in SharePoint (e.g., 'folder/data.csv' or 'data.csv')
            local_path: Local path to save the file (optional)

        Returns:
            BytesIO object containing file content if local_path is None,
            otherwise saves to local_path and returns the path
        """
        return self.download_item(file_path, local_path)

    def upload_file(self, local_file_path, sharepoint_path, file_name=None):
        """
        Upload a file to SharePoint

        Args:
            local_file_path: Path to the local file to upload
            sharepoint_path: Folder path in SharePoint (e.g., 'folder' or '' for root)
            file_name: Name for the file in SharePoint (defaults to local filename)

        Returns:
            Response from upload
        """
        if not self.drive_id:
            raise Exception("Drive ID not set. Call get_drive_id() first.")

        try:
            if file_name is None:
                file_name = os.path.basename(local_file_path)

            # Read file content
            with open(local_file_path, 'rb') as content_file:
                file_content = content_file.read()

            return self.upload_file_from_memory(file_content, sharepoint_path, file_name)

        except Exception as e:
            print(f"✗ Error uploading file: {str(e)}")
            raise

    def upload_file_from_memory(self, file_content, sharepoint_path, file_name):
        """
        Upload a file to SharePoint from memory (bytes)

        Args:
            file_content: File content as bytes
            sharepoint_path: Folder path in SharePoint (e.g., 'folder' or '' for root)
            file_name: Name for the file in SharePoint

        Returns:
            Response from upload
        """
        if not self.drive_id:
            raise Exception("Drive ID not set. Call get_drive_id() first.")

        try:
            # Build the upload URL
            if sharepoint_path:
                encoded_path = quote(f"{sharepoint_path}/{file_name}")
            else:
                encoded_path = quote(file_name)

            url = self._get_drive_item_content_url(encoded_path)

            headers = {
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/octet-stream"
            }

            # Upload file
            response = requests.put(url, headers=headers, data=file_content)
            response.raise_for_status()

            print(f"✓ File uploaded successfully: {file_name}")
            return response.json()

        except requests.exceptions.HTTPError as e:
            print(f"✗ Error uploading file: {e.response.status_code} - {e.response.text}")
            raise
        except Exception as e:
            print(f"✗ Error uploading file from memory: {str(e)}")
            raise

    def _create_item_info_from_api_response(self, item, item_type="file"):
        """
        Helper method to create ItemInfo from API response

        Args:
            item: Item dictionary from Microsoft Graph API
            item_type: Type of item ('file' or 'folder') for better path handling

        Returns:
            ItemInfo object
        """
        item_name = item.get("name", "")
        parent_path = item.get("parentReference", {}).get("path", "")

        # Extract path after the colon (removes drive prefix)
        if parent_path:
            path = parent_path.split(":")[-1] + "/" + item_name
        else:
            path = item_name

        return ItemInfo(
            name=item_name,
            path=path,
            size=item.get("size", 0),
            modified=item.get("lastModifiedDateTime", ""),
            id=item.get("id", ""),
            webUrl=item.get("webUrl", "")
        )

    def search_files_by_suffix(self, suffix, folder_path=""):
        """
        Search for files with a specific suffix/extension in SharePoint

        Args:
            suffix: File suffix/extension to search for (e.g., '.csv', '.txt', 'pdf')
            folder_path: Optional folder path to search in (e.g., 'folder' or '' for root)

        Returns:
            List of dictionaries containing file information (name, path, size, modified date)
        """
        if not self.drive_id:
            raise Exception("Drive ID not set. Call get_drive_id() first.")

        try:
            # Ensure suffix starts with a dot
            if suffix and not suffix.startswith('.'):
                suffix = f".{suffix}"

            # Build the search URL
            url = self._get_drive_children_url(folder_path)

            # Get all items in the folder
            response = requests.get(url, headers=self._get_headers())
            response.raise_for_status()

            items = response.json().get("value", [])

            # Filter files by suffix
            matching_files = []
            for item in items:
                # Check if it's a file (not a folder)
                if "file" in item:
                    file_name = item.get("name", "")
                    if file_name.endswith(suffix):
                        file_info = self._create_item_info_from_api_response(item)
                        matching_files.append(file_info)

            print(f"✓ Found {len(matching_files)} file(s) with suffix '{suffix}'")
            return matching_files

        except requests.exceptions.HTTPError as e:
            print(f"✗ Error searching files: {e.response.status_code} - {e.response.text}")
            raise
        except Exception as e:
            print(f"✗ Error searching files: {str(e)}")
            raise

    def search_files_by_suffix_recursive(self, suffix, folder_path=""):
        """
        Recursively search for files with a specific suffix/extension in SharePoint

        Args:
            suffix: File suffix/extension to search for (e.g., '.csv', '.txt', 'pdf')
            folder_path: Optional folder path to start search from (e.g., 'folder' or '' for root)

        Returns:
            List of dictionaries containing file information (name, path, size, modified date)
        """
        if not self.drive_id:
            raise Exception("Drive ID not set. Call get_drive_id() first.")

        try:
            # Ensure suffix starts with a dot
            if suffix and not suffix.startswith('.'):
                suffix = f".{suffix}"

            matching_files = []

            def search_folder_for_files(current_path):
                """Recursively search through folders for files with matching suffix"""
                # Build the URL for current folder
                url = self._get_drive_children_url(current_path)

                # Get all items in the current folder
                response = requests.get(url, headers=self._get_headers())
                response.raise_for_status()

                items = response.json().get("value", [])

                for item in items:
                    # If it's a file, check suffix
                    if "file" in item:
                        file_name = item.get("name", "")
                        if file_name.endswith(suffix):
                            file_info = self._create_item_info_from_api_response(item)
                            matching_files.append(file_info)

                    # If it's a folder, search recursively
                    elif "folder" in item:
                        folder_name = item.get("name", "")
                        new_path = f"{current_path}/{folder_name}" if current_path else folder_name
                        search_folder_for_files(new_path)

            # Start the recursive search
            search_folder_for_files(folder_path)

            print(f"✓ Found {len(matching_files)} file(s) with suffix '{suffix}' (recursive search)")
            return matching_files

        except requests.exceptions.HTTPError as e:
            print(f"✗ Error searching files: {e.response.status_code} - {e.response.text}")
            raise
        except Exception as e:
            print(f"✗ Error searching files recursively: {str(e)}")
            raise

    def search_folders_by_suffix(self, suffix, folder_path=""):
        """
        Search for folders with a specific suffix in SharePoint (non-recursive)

        Args:
            suffix: Folder suffix to search for (e.g., '.gdb', '.bundle')
            folder_path: Optional folder path to search in (e.g., 'folder' or '' for root)

        Returns:
            List of ItemInfo objects containing folder information (name, path, size, modified date, id)
        """
        if not self.drive_id:
            raise Exception("Drive ID not set. Call get_drive_id() first.")

        try:
            # Ensure suffix starts with a dot
            if suffix and not suffix.startswith('.'):
                suffix = f".{suffix}"

            # Build the search URL
            url = self._get_drive_children_url(folder_path)

            # Get all items in the folder
            response = requests.get(url, headers=self._get_headers())
            response.raise_for_status()

            items = response.json().get("value", [])

            # Filter folders by suffix
            matching_folders = []
            for item in items:
                # Check if it's a folder (not a file)
                if "folder" in item:
                    folder_name = item.get("name", "")
                    if folder_name.endswith(suffix):
                        folder_info = self._create_item_info_from_api_response(item, "folder")
                        matching_folders.append(folder_info)

            print(f"✓ Found {len(matching_folders)} folder(s) with suffix '{suffix}'")
            return matching_folders

        except requests.exceptions.HTTPError as e:
            print(f"✗ Error searching folders: {e.response.status_code} - {e.response.text}")
            raise
        except Exception as e:
            print(f"✗ Error searching folders: {str(e)}")
            raise

    def search_folders_by_suffix_recursive(self, suffix, folder_path=""):
        """
        Recursively search for folders with a specific suffix in SharePoint

        Args:
            suffix: Folder suffix to search for (e.g., '.gdb', '.bundle')
            folder_path: Optional folder path to start search from (e.g., 'folder' or '' for root)

        Returns:
            List of ItemInfo objects containing folder information (name, path, size, modified date, id)
        """
        if not self.drive_id:
            raise Exception("Drive ID not set. Call get_drive_id() first.")

        try:
            # Ensure suffix starts with a dot
            if suffix and not suffix.startswith('.'):
                suffix = f".{suffix}"

            matching_folders = []

            def search_folder_for_folders(current_path):
                """Recursively search through folders for folders with matching suffix"""
                # Build the URL for current folder
                url = self._get_drive_children_url(current_path)

                # Get all items in the current folder
                response = requests.get(url, headers=self._get_headers())
                response.raise_for_status()

                items = response.json().get("value", [])

                for item in items:
                    # If it's a folder, check suffix
                    if "folder" in item:
                        folder_name = item.get("name", "")
                        if folder_name.endswith(suffix):
                            folder_info = self._create_item_info_from_api_response(item, "folder")
                            matching_folders.append(folder_info)

                        # Continue searching in all subfolders regardless of match
                        new_path = f"{current_path}/{folder_name}" if current_path else folder_name
                        search_folder_for_folders(new_path)
                    else:
                        # If it's not a folder (it's a file), skip it
                        pass

            # Start the recursive search
            search_folder_for_folders(folder_path)

            print(f"✓ Found {len(matching_folders)} folder(s) with suffix '{suffix}' (recursive search)")
            return matching_folders

        except requests.exceptions.HTTPError as e:
            print(f"✗ Error searching folders: {e.response.status_code} - {e.response.text}")
            raise
        except Exception as e:
            print(f"✗ Error searching folders recursively: {str(e)}")
            raise

    def download_folder(self, folder_path, local_directory=None):
        """
        Recursively download an entire folder and its contents from SharePoint

        Deprecated: Use download_item() instead. This method is kept for backward compatibility.

        Args:
            folder_path: Path to the folder in SharePoint (e.g., 'folder_name' or 'parent/folder_name')
            local_directory: Local directory path to save the folder (defaults to folder name in current directory)

        Returns:
            Path to the downloaded folder
        """
        return self.download_item(folder_path, local_directory)

    def delete_item(self, item_path):
        """
        Delete a file or folder from SharePoint

        Args:
            item_path: Path to the file or folder in SharePoint (e.g., 'folder/file.txt', 'folder_name')

        Returns:
            True if deletion was successful

        Warning:
            If deleting a folder, this operation is permanent and will delete all files and subfolders within.
        """
        if not self.drive_id:
            raise Exception("Drive ID not set. Call get_drive_id() first.")

        try:
            # Encode the item path
            encoded_path = quote(item_path)

            # Build the delete URL
            url = self._get_drive_item_url(encoded_path)

            # Delete item
            response = requests.delete(url, headers=self._get_headers())
            response.raise_for_status()

            print(f"✓ Item deleted successfully: {item_path}")
            return True

        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                print(f"✗ Item not found: {item_path}")
            else:
                print(f"✗ Error deleting item: {e.response.status_code} - {e.response.text}")
            raise
        except Exception as e:
            print(f"✗ Error deleting item: {str(e)}")
            raise

    def delete_file(self, file_path):
        """
        Delete a file from SharePoint

        Deprecated: Use delete_item() instead. This method is kept for backward compatibility.

        Args:
            file_path: Path to the file in SharePoint (e.g., 'folder/file.txt' or 'file.txt')

        Returns:
            True if deletion was successful
        """
        return self.delete_item(file_path)

    def delete_folder(self, folder_path):
        """
        Delete a folder and all its contents from SharePoint

        Deprecated: Use delete_item() instead. This method is kept for backward compatibility.

        Args:
            folder_path: Path to the folder in SharePoint (e.g., 'folder_name' or 'parent/folder_name')

        Returns:
            True if deletion was successful

        Warning:
            This operation is permanent and will delete all files and subfolders within.
        """
        return self.delete_item(folder_path)

    def move_item(self, item_path, destination_folder_path):
        """
        Move a file or folder to a different location in SharePoint

        Args:
            item_path: Current path to the file or folder (e.g., 'current_folder/file.txt' or 'folder_name')
            destination_folder_path: Destination folder path (e.g., 'new_folder' or 'archive/new_folder')

        Returns:
            True if move was successful
        """
        if not self.drive_id:
            raise Exception("Drive ID not set. Call get_drive_id() first.")

        try:
            # Get the item ID
            item_id = self._get_item_id_by_path(item_path)

            # Get the destination folder ID
            destination_folder_id = self._get_item_id_by_path(destination_folder_path)

            # Build the move URL
            url = f"{self.GRAPH_API_BASE}/sites/{self.site_id}/drives/{self.drive_id}/items/{item_id}"

            # Prepare the move request body
            move_body = {
                "parentReference": {
                    "id": destination_folder_id
                }
            }

            # Move item
            response = requests.patch(url, headers=self._get_headers(), json=move_body)
            response.raise_for_status()

            print(f"✓ Item moved successfully: {item_path} → {destination_folder_path}")
            return True

        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                print("✗ Item or destination folder not found")
            else:
                print(f"✗ Error moving item: {e.response.status_code} - {e.response.text}")
            raise
        except Exception as e:
            print(f"✗ Error moving item: {str(e)}")
            raise

    def move_file(self, file_path, destination_folder_path):
        """
        Move a file to a different folder in SharePoint

        Deprecated: Use move_item() instead. This method is kept for backward compatibility.

        Args:
            file_path: Current path to the file (e.g., 'current_folder/file.txt')
            destination_folder_path: Destination folder path (e.g., 'new_folder' or 'archive/new_folder')

        Returns:
            True if move was successful
        """
        return self.move_item(file_path, destination_folder_path)

    def move_folder(self, folder_path, destination_parent_folder_path):
        """
        Move a folder to a different location in SharePoint

        Deprecated: Use move_item() instead. This method is kept for backward compatibility.

        Args:
            folder_path: Current path to the folder (e.g., 'current_location/folder_name')
            destination_parent_folder_path: Destination parent folder path (e.g., 'archive' or 'backup/old_folders')

        Returns:
            True if move was successful
        """
        return self.move_item(folder_path, destination_parent_folder_path)