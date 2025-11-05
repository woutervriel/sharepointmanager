"""
Tests for SharePointManager initialization and authentication
"""
import pytest
from unittest.mock import Mock, patch
from sharepointmanager.sharepoint import SharePointManager


class TestInitialization:
    """Tests for SharePointManager initialization"""

    def test_manager_initialization(self, config):
        """Test basic manager initialization"""
        print(config)
        with patch('sharepointmanager.sharepoint.msal.ConfidentialClientApplication') as mock_auth:
            # Configure the mock to return a proper access token
            mock_app = Mock()
            mock_app.acquire_token_for_client.return_value = {
                "access_token": "test-token"
            }
            mock_auth.return_value = mock_app

            manager = SharePointManager(
                tenant_id=config["TENANT_ID"],
                client_id=config["CLIENT_ID"],
                client_secret=config["CLIENT_SECRET"],
                site_name=config["SITE_NAME"]
            )

            assert manager.tenant_id == config["TENANT_ID"]
            assert manager.client_id == config["CLIENT_ID"]
            assert manager.client_secret == config["CLIENT_SECRET"]

    def test_site_name_formatting_without_domain(self):
        """Test site name is auto-formatted with domain"""
        with patch('sharepointmanager.sharepoint.msal.ConfidentialClientApplication') as mock_auth:
            mock_app = Mock()
            mock_app.acquire_token_for_client.return_value = {"access_token": "test-token"}
            mock_auth.return_value = mock_app

            manager = SharePointManager(
                tenant_id="test-tenant",
                client_id="test-client",
                client_secret="test-secret",
                site_name="mysite"
            )

            assert manager.site_name == "mysite.sharepoint.com"

    def test_site_name_formatting_with_domain(self):
        """Test site name with domain is kept as-is"""
        with patch('sharepointmanager.sharepoint.msal.ConfidentialClientApplication') as mock_auth:
            mock_app = Mock()
            mock_app.acquire_token_for_client.return_value = {"access_token": "test-token"}
            mock_auth.return_value = mock_app

            manager = SharePointManager(
                tenant_id="test-tenant",
                client_id="test-client",
                client_secret="test-secret",
                site_name="mysite.sharepoint.com"
            )

            assert manager.site_name == "mysite.sharepoint.com"

    def test_initial_ids_are_none(self, config):
        """Test that IDs are initially None"""
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

            assert manager.site_id is None
            assert manager.drive_id is None

    def test_access_token_set_after_auth(self, config):
        """Test that access token is set after authentication"""
        with patch('sharepointmanager.sharepoint.msal.ConfidentialClientApplication') as mock_auth:
            mock_app = Mock()
            mock_app.acquire_token_for_client.return_value = {
                "access_token": "test-token"
            }
            mock_auth.return_value = mock_app

            manager = SharePointManager(
                tenant_id=config["TENANT_ID"],
                client_id=config["CLIENT_ID"],
                client_secret=config["CLIENT_SECRET"],
                site_name=config["SITE_NAME"]
            )

            assert manager.access_token == "test-token"


class TestGetSiteId:
    """Tests for get_site_id method"""

    def test_get_site_id_success(self, mock_sharepoint_manager, mock_requests_get):
        """Test getting site ID successfully"""
        response = Mock()
        response.json.return_value = {"id": "site-123"}
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        result = mock_sharepoint_manager.get_site_id()

        assert result == "site-123"
        assert mock_sharepoint_manager.site_id == "site-123"

    def test_get_site_id_with_path(self, mock_sharepoint_manager, mock_requests_get):
        """Test getting site ID with site path"""
        response = Mock()
        response.json.return_value = {"id": "site-456"}
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        result = mock_sharepoint_manager.get_site_id("/sites/mysite")

        assert result == "site-456"


class TestGetDriveId:
    """Tests for get_drive_id method"""

    def test_get_drive_id_success(self, mock_sharepoint_manager, mock_requests_get):
        """Test getting drive ID successfully"""
        response = Mock()
        response.json.return_value = {
            "value": [
                {"name": "Documents", "id": "drive-123"}
            ]
        }
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        result = mock_sharepoint_manager.get_drive_id("Documents")

        assert result == "drive-123"
        assert mock_sharepoint_manager.drive_id == "drive-123"

    def test_get_drive_id_default_name(self, mock_sharepoint_manager, mock_requests_get):
        """Test getting drive ID with default name"""
        response = Mock()
        response.json.return_value = {
            "value": [
                {"name": "Documenten", "id": "drive-456"}
            ]
        }
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        result = mock_sharepoint_manager.get_drive_id()

        assert result == "drive-456"

    def test_get_drive_id_not_found(self, mock_sharepoint_manager, mock_requests_get):
        """Test getting non-existent drive"""
        response = Mock()
        response.json.return_value = {"value": []}
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        with pytest.raises(Exception, match="No drives found"):
            mock_sharepoint_manager.get_drive_id("NonExistent")

    def test_get_drive_id_without_site_id(self, config):
        """Test get_drive_id raises exception without site_id"""
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
            manager.site_id = None

            with pytest.raises(Exception, match="Site ID not set"):
                manager.get_drive_id()

    def test_get_drive_id_fallback_to_first_drive(self, mock_sharepoint_manager, mock_requests_get):
        """Test fallback to first drive if named drive not found"""
        response = Mock()
        response.json.return_value = {
            "value": [
                {"name": "Documenten", "id": "drive-789"}
            ]
        }
        response.raise_for_status.return_value = None

        mock_requests_get.return_value = response

        result = mock_sharepoint_manager.get_drive_id("NonExistent")

        assert result == "drive-789"


class TestHeaders:
    """Tests for header generation"""

    def test_get_headers(self, mock_sharepoint_manager):
        """Test header generation"""
        headers = mock_sharepoint_manager._get_headers()

        assert "Authorization" in headers
        assert headers["Authorization"] == "Bearer test-token"
        assert headers["Content-Type"] == "application/json"

    def test_headers_contain_access_token(self, mock_sharepoint_manager):
        """Test headers contain correct access token"""
        mock_sharepoint_manager.access_token = "custom-token"
        headers = mock_sharepoint_manager._get_headers()

        assert "custom-token" in headers["Authorization"]


class TestAuthentication:
    """Tests for authentication process"""

    def test_authentication_success(self, config):
        """Test successful authentication"""
        with patch('sharepointmanager.sharepoint.msal.ConfidentialClientApplication') as mock_auth:
            mock_app = Mock()
            mock_app.acquire_token_for_client.return_value = {
                "access_token": "auth-token"
            }
            mock_auth.return_value = mock_app

            manager = SharePointManager(
                tenant_id=config["TENANT_ID"],
                client_id=config["CLIENT_ID"],
                client_secret=config["CLIENT_SECRET"],
                site_name=config["SITE_NAME"]
            )

            assert manager.access_token is not None

    def test_authentication_failure(self, config):
        """Test authentication failure"""
        with patch('sharepointmanager.sharepoint.msal.ConfidentialClientApplication') as mock_auth:
            mock_app = Mock()
            mock_app.acquire_token_for_client.return_value = {
                "error": "invalid_client",
                "error_description": "Client authentication failed"
            }
            mock_auth.return_value = mock_app

            with pytest.raises(Exception, match="Authentication failed"):
                manager = SharePointManager(
                    tenant_id=config["TENANT_ID"],
                    client_id=config["CLIENT_ID"],
                    client_secret=config["CLIENT_SECRET"],
                    site_name=config["SITE_NAME"]
                )

    def test_authentication_with_msal_app(self, config):
        """Test that MSAL app is created correctly"""
        with patch('sharepointmanager.sharepoint.msal.ConfidentialClientApplication') as mock_auth:
            mock_app = Mock()
            mock_app.acquire_token_for_client.return_value = {
                "access_token": "token"
            }
            mock_auth.return_value = mock_app

            manager = SharePointManager(
                tenant_id=config["TENANT_ID"],
                client_id=config["CLIENT_ID"],
                client_secret=config["CLIENT_SECRET"],
                site_name=config["SITE_NAME"]
            )

            # Verify MSAL was called with correct parameters
            mock_auth.assert_called_once()
            call_args = mock_auth.call_args
            assert config["CLIENT_ID"] in call_args[1].values() or call_args[0][0] == config["CLIENT_ID"]
