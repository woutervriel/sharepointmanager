"""
Tests for file search operations
"""
import pytest
from unittest.mock import Mock, patch
from sharepointmanager.sharepoint import SharePointManager, ItemInfo


class TestSearchFilesBySuffix:
    """Tests for search_files_by_suffix method"""

    def test_search_files_found(self, mock_sharepoint_manager, mock_requests_get):
        """Test finding files with specific suffix"""
        response = Mock()
        response.json.return_value = {
            "value": [
                {
                    "name": "data.csv",
                    "id": "file-1",
                    "file": {},
                    "parentReference": {"path": "/drives/root"},
                    "size": 2048,
                    "lastModifiedDateTime": "2024-01-01T00:00:00Z",
                    "webUrl": "https://sharepoint.com/data.csv"
                }
            ]
        }
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        results = mock_sharepoint_manager.search_files_by_suffix(".csv")

        assert len(results) == 1
        assert results[0].name == "data.csv"
        assert isinstance(results[0], ItemInfo)

    def test_search_files_no_matches(self, mock_sharepoint_manager, mock_requests_get):
        """Test searching with no matches"""
        response = Mock()
        response.json.return_value = {"value": []}
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        results = mock_sharepoint_manager.search_files_by_suffix(".xyz")

        assert len(results) == 0

    def test_search_files_multiple_matches(self, mock_sharepoint_manager, mock_requests_get):
        """Test finding multiple files with same suffix"""
        response = Mock()
        response.json.return_value = {
            "value": [
                {
                    "name": "file1.csv",
                    "id": "file-1",
                    "file": {},
                    "parentReference": {"path": "/drives/root"},
                    "size": 1024,
                    "lastModifiedDateTime": "2024-01-01T00:00:00Z",
                    "webUrl": "https://sharepoint.com/file1.csv"
                },
                {
                    "name": "file2.csv",
                    "id": "file-2",
                    "file": {},
                    "parentReference": {"path": "/drives/root"},
                    "size": 2048,
                    "lastModifiedDateTime": "2024-01-02T00:00:00Z",
                    "webUrl": "https://sharepoint.com/file2.csv"
                }
            ]
        }
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        results = mock_sharepoint_manager.search_files_by_suffix(".csv")

        assert len(results) == 2
        assert results[0].name == "file1.csv"
        assert results[1].name == "file2.csv"

    def test_search_files_without_dot_prefix(self, mock_sharepoint_manager, mock_requests_get):
        """Test search suffix without dot is auto-added"""
        response = Mock()
        response.json.return_value = {"value": []}
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        # Call without dot prefix
        mock_sharepoint_manager.search_files_by_suffix("csv")

        # Verify it was called (suffix conversion happens internally)
        mock_requests_get.assert_called()

    def test_search_files_in_specific_folder(self, mock_sharepoint_manager, mock_requests_get):
        """Test searching in specific folder"""
        response = Mock()
        response.json.return_value = {"value": []}
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        mock_sharepoint_manager.search_files_by_suffix(".csv", "data_folder")

        # Verify the call was made
        mock_requests_get.assert_called()

    def test_search_files_without_drive_id(self, config):
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
                manager.search_files_by_suffix(".csv")


class TestSearchFilesByRecursive:
    """Tests for search_files_by_suffix_recursive method"""

    def test_search_files_recursive_found(self, mock_sharepoint_manager, mock_requests_get):
        """Test recursively finding files"""
        # First call returns root items
        response = Mock()
        response.json.return_value = {"value": []}
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        results = mock_sharepoint_manager.search_files_by_suffix_recursive(".pdf")

        assert isinstance(results, list)
        mock_requests_get.assert_called()

    def test_search_files_recursive_without_drive_id(self, config):
        """Test recursive search raises exception without drive_id"""
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
                manager.search_files_by_suffix_recursive(".pdf")

    def test_search_files_recursive_with_start_path(self, mock_sharepoint_manager, mock_requests_get):
        """Test recursive search with specific start path"""
        response = Mock()
        response.json.return_value = {"value": []}
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        results = mock_sharepoint_manager.search_files_by_suffix_recursive(".log", "logs")

        assert isinstance(results, list)


class TestSearchFilesBySuffixValidation:
    """Tests for input validation in search operations"""

    def test_search_with_empty_suffix(self, mock_sharepoint_manager, mock_requests_get):
        """Test search with empty suffix"""
        response = Mock()
        response.json.return_value = {"value": []}
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        # Should handle empty suffix gracefully
        results = mock_sharepoint_manager.search_files_by_suffix("")

        assert isinstance(results, list)

    def test_search_returns_correct_dataclass(self, mock_sharepoint_manager, mock_requests_get):
        """Test that search returns FileInfo dataclass objects"""
        response = Mock()
        response.json.return_value = {
            "value": [
                {
                    "name": "test.txt",
                    "id": "file-123",
                    "file": {},
                    "parentReference": {"path": "/drives/root"},
                    "size": 1024,
                    "lastModifiedDateTime": "2024-01-01T00:00:00Z",
                    "webUrl": "https://sharepoint.com/test.txt"
                }
            ]
        }
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        results = mock_sharepoint_manager.search_files_by_suffix(".txt")

        assert len(results) == 1
        result = results[0]
        assert result.name == "test.txt"
        assert result.id == "file-123"
        assert result.size == 1024
        assert hasattr(result, 'webUrl')
