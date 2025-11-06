"""
Tests for folder operations (download, delete, move, search)
"""
import pytest
from unittest.mock import Mock, patch
from requests.exceptions import HTTPError
from sharepointer.sharepoint import SharePointManager, ItemInfo


class TestSearchFoldersByPrefix:
    """Tests for search_folders_by_suffix_recursive method"""

    def test_search_folders_found(self, mock_sharepoint_manager, mock_requests_get):
        """Test finding folders with specific suffix"""
        # First response contains the matching folder
        response1 = Mock()
        response1.json.return_value = {
            "value": [
                {
                    "name": "database.gdb",
                    "id": "folder-1",
                    "folder": {},
                    "parentReference": {"path": "/drives/root"},
                    "size": 5120,
                    "lastModifiedDateTime": "2024-01-01T00:00:00Z",
                    "webUrl": "https://sharepoint.com/database.gdb"
                }
            ]
        }
        response1.raise_for_status.return_value = None

        # Second response (when searching inside the folder) is empty
        response2 = Mock()
        response2.json.return_value = {"value": []}
        response2.raise_for_status.return_value = None

        # Return different responses for each call
        mock_requests_get.side_effect = [response1, response2]

        results = mock_sharepoint_manager.search_folders_by_suffix_recursive(".gdb")

        assert len(results) > 0
        assert results[0].name == "database.gdb"

    def test_search_folders_no_matches(self, mock_sharepoint_manager, mock_requests_get):
        """Test searching with no matches"""
        response = Mock()
        response.json.return_value = {"value": []}
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        results = mock_sharepoint_manager.search_folders_by_suffix_recursive(".xyz")

        assert len(results) == 0

    def test_search_folders_without_dot_prefix(self, mock_sharepoint_manager, mock_requests_get):
        """Test search suffix without dot is auto-added"""
        response = Mock()
        response.json.return_value = {"value": []}
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        # Call without dot prefix
        mock_sharepoint_manager.search_folders_by_suffix_recursive("gdb")

        # Verify the suffix was converted
        mock_requests_get.assert_called()

    def test_search_folders_without_drive_id(self, config):
        """Test search raises exception without drive_id"""
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
                manager.search_folders_by_suffix_recursive(".gdb")


class TestSearchFoldersBySuffix:
    """Tests for search_folders_by_suffix method (non-recursive)"""

    def test_search_folders_found(self, mock_sharepoint_manager, mock_requests_get):
        """Test finding folders with specific suffix in a single folder"""
        response = Mock()
        response.json.return_value = {
            "value": [
                {
                    "name": "data.gdb",
                    "id": "folder-1",
                    "folder": {},
                    "parentReference": {"path": "/drives/root"},
                    "size": 2048,
                    "lastModifiedDateTime": "2024-01-01T00:00:00Z",
                    "webUrl": "https://sharepoint.com/data.gdb"
                },
                {
                    "name": "archive.gdb",
                    "id": "folder-2",
                    "folder": {},
                    "parentReference": {"path": "/drives/root"},
                    "size": 4096,
                    "lastModifiedDateTime": "2024-01-02T00:00:00Z",
                    "webUrl": "https://sharepoint.com/archive.gdb"
                },
                {
                    "name": "other_folder",
                    "id": "folder-3",
                    "folder": {},
                    "parentReference": {"path": "/drives/root"},
                    "size": 1024,
                    "lastModifiedDateTime": "2024-01-03T00:00:00Z",
                    "webUrl": "https://sharepoint.com/other_folder"
                }
            ]
        }
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        results = mock_sharepoint_manager.search_folders_by_suffix(".gdb")

        assert len(results) == 2
        assert results[0].name == "data.gdb"
        assert results[1].name == "archive.gdb"
        assert isinstance(results[0], ItemInfo)

    def test_search_folders_no_matches(self, mock_sharepoint_manager, mock_requests_get):
        """Test searching with no matching folders"""
        response = Mock()
        response.json.return_value = {
            "value": [
                {
                    "name": "regular_folder",
                    "id": "folder-1",
                    "folder": {},
                    "parentReference": {"path": "/drives/root"},
                    "size": 1024,
                    "lastModifiedDateTime": "2024-01-01T00:00:00Z",
                    "webUrl": "https://sharepoint.com/regular_folder"
                }
            ]
        }
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        results = mock_sharepoint_manager.search_folders_by_suffix(".gdb")

        assert len(results) == 0

    def test_search_folders_without_dot_prefix(self, mock_sharepoint_manager, mock_requests_get):
        """Test search suffix without dot is auto-added"""
        response = Mock()
        response.json.return_value = {
            "value": [
                {
                    "name": "test.gdb",
                    "id": "folder-1",
                    "folder": {},
                    "parentReference": {"path": "/drives/root"},
                    "size": 1024,
                    "lastModifiedDateTime": "2024-01-01T00:00:00Z",
                    "webUrl": "https://sharepoint.com/test.gdb"
                }
            ]
        }
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        # Call without dot prefix
        results = mock_sharepoint_manager.search_folders_by_suffix("gdb")

        assert len(results) == 1
        assert results[0].name == "test.gdb"

    def test_search_folders_in_specific_folder(self, mock_sharepoint_manager, mock_requests_get):
        """Test searching in a specific folder path"""
        response = Mock()
        response.json.return_value = {
            "value": [
                {
                    "name": "project.gdb",
                    "id": "folder-1",
                    "folder": {},
                    "parentReference": {"path": "/drives/root/Data"},
                    "size": 3072,
                    "lastModifiedDateTime": "2024-01-01T00:00:00Z",
                    "webUrl": "https://sharepoint.com/Data/project.gdb"
                }
            ]
        }
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        results = mock_sharepoint_manager.search_folders_by_suffix(".gdb", "Data")

        assert len(results) == 1
        assert results[0].name == "project.gdb"
        mock_requests_get.assert_called_once()

    def test_search_folders_only_returns_folders_not_files(self, mock_sharepoint_manager, mock_requests_get):
        """Test that search only returns folders, not files"""
        response = Mock()
        response.json.return_value = {
            "value": [
                {
                    "name": "folder.gdb",
                    "id": "folder-1",
                    "folder": {},
                    "parentReference": {"path": "/drives/root"},
                    "size": 2048,
                    "lastModifiedDateTime": "2024-01-01T00:00:00Z",
                    "webUrl": "https://sharepoint.com/folder.gdb"
                },
                {
                    "name": "file.gdb",
                    "id": "file-1",
                    "file": {},
                    "parentReference": {"path": "/drives/root"},
                    "size": 1024,
                    "lastModifiedDateTime": "2024-01-02T00:00:00Z",
                    "webUrl": "https://sharepoint.com/file.gdb"
                }
            ]
        }
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        results = mock_sharepoint_manager.search_folders_by_suffix(".gdb")

        # Should only return the folder, not the file
        assert len(results) == 1
        assert results[0].name == "folder.gdb"

    def test_search_folders_without_drive_id(self, config):
        """Test search raises exception without drive_id"""
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
                manager.search_folders_by_suffix(".gdb")


class TestDeleteFolder:
    """Tests for delete_folder method"""

    def test_delete_folder_success(self, mock_sharepoint_manager, mock_requests_delete):
        """Test successful folder deletion"""
        response = Mock()
        response.raise_for_status.return_value = None

        mock_requests_delete.return_value = response

        result = mock_sharepoint_manager.delete_folder("old_folder")

        assert result is True
        mock_requests_delete.assert_called_once()

    def test_delete_folder_not_found(self, mock_sharepoint_manager, mock_requests_delete):
        """Test deleting non-existent folder"""
        response = Mock()
        response.status_code = 404
        response.text = "Not Found"
        response.raise_for_status.side_effect = HTTPError(response=response)

        mock_requests_delete.return_value = response

        with pytest.raises(HTTPError):
            mock_sharepoint_manager.delete_folder("nonexistent_folder")

    def test_delete_folder_without_drive_id(self, config):
        """Test delete_folder raises exception without drive_id"""
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
                manager.delete_folder("test_folder")


class TestMoveFolder:
    """Tests for move_folder method"""

    def test_move_folder_success(self, mock_sharepoint_manager, mock_requests_get, mock_requests_patch):
        """Test successful folder move"""
        # Mock getting folder ID
        folder_response = Mock()
        folder_response.json.return_value = {"id": "folder-123"}
        folder_response.raise_for_status.return_value = None

        # Mock getting destination parent folder ID
        dest_response = Mock()
        dest_response.json.return_value = {"id": "folder-456"}
        dest_response.raise_for_status.return_value = None

        mock_requests_get.side_effect = [folder_response, dest_response]

        # Mock patch response
        patch_response = Mock()
        patch_response.raise_for_status.return_value = None

        mock_requests_patch.return_value = patch_response

        result = mock_sharepoint_manager.move_folder("active/old_project", "archive")

        assert result is True
        mock_requests_patch.assert_called_once()

    def test_move_folder_not_found(self, mock_sharepoint_manager, mock_requests_get):
        """Test moving non-existent folder"""
        response = Mock()
        response.status_code = 404
        response.text = "Not Found"
        response.raise_for_status.side_effect = HTTPError(response=response)

        mock_requests_get.return_value = response

        with pytest.raises(Exception):
            mock_sharepoint_manager.move_folder("nonexistent", "archive")

    def test_move_folder_without_drive_id(self, config):
        """Test move_folder raises exception without drive_id"""
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
                manager.move_folder("test_folder", "archive")


class TestDownloadFolder:
    """Tests for download_folder method"""

    def test_download_folder_path_creation(self, mock_sharepoint_manager, mock_requests_get):
        """Test folder download path creation"""
        # Mock metadata response
        metadata_response = Mock()
        metadata_response.json.return_value = {"folder": {}}
        metadata_response.raise_for_status.return_value = None

        # Mock folder contents response
        contents_response = Mock()
        contents_response.json.return_value = {"value": []}
        contents_response.raise_for_status.return_value = None

        mock_requests_get.side_effect = [metadata_response, contents_response]

        with patch('os.makedirs'):
            with patch.object(mock_sharepoint_manager, '_get_drive_children_url', return_value="url"):
                result = mock_sharepoint_manager.download_folder("test_folder")

                assert result == "test_folder"

    def test_download_folder_without_drive_id(self, config):
        """Test download_folder raises exception without drive_id"""
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
                manager.download_folder("test_folder")

    def test_download_folder_custom_path(self, mock_sharepoint_manager, mock_requests_get):
        """Test folder download with custom local path"""
        # Mock metadata response
        metadata_response = Mock()
        metadata_response.json.return_value = {"folder": {}}
        metadata_response.raise_for_status.return_value = None

        # Mock folder contents response
        contents_response = Mock()
        contents_response.json.return_value = {"value": []}
        contents_response.raise_for_status.return_value = None

        mock_requests_get.side_effect = [metadata_response, contents_response]

        with patch('os.makedirs'):
            with patch.object(mock_sharepoint_manager, '_get_drive_children_url', return_value="url"):
                result = mock_sharepoint_manager.download_folder("test_folder", "/custom/path")

                assert result == "/custom/path"
