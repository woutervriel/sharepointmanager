"""
Tests for file operations (download, upload, delete, move)
"""
import pytest
from unittest.mock import Mock, patch
from io import BytesIO
from requests.exceptions import HTTPError
from sharepointmanager.sharepoint import SharePointManager


class TestDownloadFile:
    """Tests for download_file method"""

    def test_download_file_to_memory(self, mock_sharepoint_manager, mock_requests_get):
        """Test downloading file to memory"""
        # Mock the metadata response
        metadata_response = Mock()
        metadata_response.json.return_value = {
            "id": "file-123",
            "file": {},
            "@microsoft.graph.downloadUrl": "https://download.sharepoint.com/file"
        }
        metadata_response.raise_for_status.return_value = None

        # Mock the download response
        download_response = Mock()
        download_response.content = b"file content"
        download_response.raise_for_status.return_value = None

        mock_requests_get.side_effect = [metadata_response, download_response]

        result = mock_sharepoint_manager.download_file("test.txt")

        assert isinstance(result, BytesIO)
        assert result.getvalue() == b"file content"

    def test_download_file_to_local_path(self, mock_sharepoint_manager, mock_requests_get):
        """Test downloading file to local path"""
        # Mock the metadata response
        metadata_response = Mock()
        metadata_response.json.return_value = {
            "file": {},
            "@microsoft.graph.downloadUrl": "https://download.sharepoint.com/file"
        }
        metadata_response.raise_for_status.return_value = None

        # Mock the download response
        download_response = Mock()
        download_response.content = b"file content"
        download_response.raise_for_status.return_value = None

        mock_requests_get.side_effect = [metadata_response, download_response]

        with patch('builtins.open', create=True) as mock_file:
            result = mock_sharepoint_manager.download_file("test.txt", "/tmp/test.txt")

            assert result == "/tmp/test.txt"
            mock_file.assert_called_once_with("/tmp/test.txt", 'wb')

    def test_download_file_not_found(self, mock_sharepoint_manager, mock_requests_get):
        """Test downloading non-existent file"""
        response = Mock()
        response.status_code = 404
        response.text = "Not Found"
        response.raise_for_status.side_effect = HTTPError(response=response)

        mock_requests_get.return_value = response

        with pytest.raises(HTTPError):
            mock_sharepoint_manager.download_file("nonexistent.txt")

    def test_download_file_without_drive_id(self, config):
        """Test download_file raises exception without drive_id"""
        with patch('sharepointmanager.sharepoint.msal.ConfidentialClientApplication') as mock_auth:
            mock_app = Mock()
            mock_app.acquire_token_for_client.return_value = {"access_token": "test-token"}
            mock_auth.return_value = mock_app

            manager = SharePointManager(
                tenant_id=config["TENANT_ID"],
                client_id=config["CLIENT_ID"],
                client_secret=config["CLIENT_SECRET"],
                site_name=config["SITE_NAME"]
            )
            manager.access_token = "token"
            manager.site_id = "site-id"
            manager.drive_id = None

            with pytest.raises(Exception, match="Drive ID not set"):
                manager.download_file("test.txt")


class TestUploadFile:
    """Tests for upload_file and upload_file_from_memory methods"""

    def test_upload_file_from_memory(self, mock_sharepoint_manager, mock_requests_put):
        """Test uploading file from memory"""
        response = Mock()
        response.json.return_value = {"id": "file-123", "name": "test.txt"}
        response.raise_for_status.return_value = None

        mock_requests_put.return_value = response

        file_content = b"test content"
        result = mock_sharepoint_manager.upload_file_from_memory(file_content, "folder", "test.txt")

        assert result["id"] == "file-123"
        mock_requests_put.assert_called_once()

    def test_upload_file_to_root(self, mock_sharepoint_manager, mock_requests_put):
        """Test uploading file to root folder"""
        response = Mock()
        response.json.return_value = {"id": "file-456"}
        response.raise_for_status.return_value = None

        mock_requests_put.return_value = response

        file_content = b"test content"
        mock_sharepoint_manager.upload_file_from_memory(file_content, "", "test.txt")

        # Check that the URL contains the file name
        call_args = mock_requests_put.call_args
        assert "test.txt" in call_args[0][0] or "test.txt" in str(call_args)

    def test_upload_file_from_path(self, mock_sharepoint_manager, mock_requests_put):
        """Test uploading file from local path"""
        response = Mock()
        response.json.return_value = {"id": "file-789"}
        response.raise_for_status.return_value = None

        mock_requests_put.return_value = response

        with patch('builtins.open', create=True) as mock_file:
            mock_file.return_value.__enter__.return_value.read.return_value = b"content"
            mock_sharepoint_manager.upload_file("/local/path/test.txt", "folder")

            mock_requests_put.assert_called_once()


class TestDeleteFile:
    """Tests for delete_file method"""

    def test_delete_file_success(self, mock_sharepoint_manager, mock_requests_delete):
        """Test successful file deletion"""
        response = Mock()
        response.raise_for_status.return_value = None

        mock_requests_delete.return_value = response

        result = mock_sharepoint_manager.delete_file("test.txt")

        assert result is True
        mock_requests_delete.assert_called_once()

    def test_delete_file_not_found(self, mock_sharepoint_manager, mock_requests_delete):
        """Test deleting non-existent file"""
        response = Mock()
        response.status_code = 404
        response.text = "Not Found"
        response.raise_for_status.side_effect = HTTPError(response=response)

        mock_requests_delete.return_value = response

        with pytest.raises(HTTPError):
            mock_sharepoint_manager.delete_file("nonexistent.txt")

    def test_delete_file_without_drive_id(self, config):
        """Test delete_file raises exception without drive_id"""
        with patch('sharepointmanager.sharepoint.msal.ConfidentialClientApplication') as mock_auth:
            mock_app = Mock()
            mock_app.acquire_token_for_client.return_value = {"access_token": "test-token"}
            mock_auth.return_value = mock_app

            manager = SharePointManager(
                tenant_id=config["TENANT_ID"],
                client_id=config["CLIENT_ID"],
                client_secret=config["CLIENT_SECRET"],
                site_name=config["SITE_NAME"]
            )
            manager.access_token = "token"
            manager.site_id = "site-id"
            manager.drive_id = None

            with pytest.raises(Exception, match="Drive ID not set"):
                manager.delete_file("test.txt")


class TestMoveFile:
    """Tests for move_file method"""

    def test_move_file_success(self, mock_sharepoint_manager, mock_requests_get, mock_requests_patch):
        """Test successful file move"""
        # Mock getting file ID
        file_response = Mock()
        file_response.json.return_value = {"id": "file-123"}
        file_response.raise_for_status.return_value = None

        # Mock getting destination folder ID
        folder_response = Mock()
        folder_response.json.return_value = {"id": "folder-456"}
        folder_response.raise_for_status.return_value = None

        mock_requests_get.side_effect = [file_response, folder_response]

        # Mock patch response
        patch_response = Mock()
        patch_response.raise_for_status.return_value = None

        mock_requests_patch.return_value = patch_response

        result = mock_sharepoint_manager.move_file("test.txt", "archive")

        assert result is True
        mock_requests_patch.assert_called_once()

    def test_move_file_not_found(self, mock_sharepoint_manager, mock_requests_get):
        """Test moving non-existent file"""
        response = Mock()
        response.status_code = 404
        response.text = "Not Found"
        response.raise_for_status.side_effect = HTTPError(response=response)

        mock_requests_get.return_value = response

        with pytest.raises(Exception):
            mock_sharepoint_manager.move_file("nonexistent.txt", "archive")

    def test_move_file_without_drive_id(self, config):
        """Test move_file raises exception without drive_id"""
        with patch('sharepointmanager.sharepoint.msal.ConfidentialClientApplication') as mock_auth:
            mock_app = Mock()
            mock_app.acquire_token_for_client.return_value = {"access_token": "test-token"}
            mock_auth.return_value = mock_app

            manager = SharePointManager(
                tenant_id=config["TENANT_ID"],
                client_id=config["CLIENT_ID"],
                client_secret=config["CLIENT_SECRET"],
                site_name=config["SITE_NAME"]
            )
            manager.access_token = "token"
            manager.site_id = "site-id"
            manager.drive_id = None

            with pytest.raises(Exception, match="Drive ID not set"):
                manager.move_file("test.txt", "archive")
