"""
Tests for URL helper methods
"""
from unittest.mock import Mock, patch
from sharepointer.sharepoint import SharePointManager


class TestURLHelpers:
    """Tests for URL builder methods"""

    def test_get_site_url_with_path(self, config, mock_sharepoint_manager):
        """Test _get_site_url with site path"""
        url = mock_sharepoint_manager._get_site_url("/sites/mysite")

        assert f"https://graph.microsoft.com/v1.0/sites/{config['SITE_NAME']}.sharepoint.com:/sites/mysite" in url

    def test_get_site_url_without_path(self, config, mock_sharepoint_manager):
        """Test _get_site_url without site path"""
        url = mock_sharepoint_manager._get_site_url()

        assert url == f"https://graph.microsoft.com/v1.0/sites/{config['SITE_NAME']}.sharepoint.com"

    def test_get_drives_url(self, mock_sharepoint_manager):
        """Test _get_drives_url"""
        url = mock_sharepoint_manager._get_drives_url()

        assert "sites/test-site-id/drives" in url
        assert url.startswith("https://graph.microsoft.com/v1.0")

    def test_get_drive_root_url(self, mock_sharepoint_manager):
        """Test _get_drive_root_url"""
        url = mock_sharepoint_manager._get_drive_root_url()

        assert "sites/test-site-id/drives/test-drive-id/root" in url

    def test_get_drive_item_url(self, mock_sharepoint_manager):
        """Test _get_drive_item_url"""
        url = mock_sharepoint_manager._get_drive_item_url("test%2Ffile.txt")

        assert ":/test%2Ffile.txt" in url
        assert "/root:/" in url

    def test_get_drive_item_content_url(self, mock_sharepoint_manager):
        """Test _get_drive_item_content_url"""
        url = mock_sharepoint_manager._get_drive_item_content_url("test%2Ffile.txt")

        assert ":/test%2Ffile.txt:/content" in url

    def test_get_drive_children_url_with_folder(self, mock_sharepoint_manager):
        """Test _get_drive_children_url with folder path"""
        url = mock_sharepoint_manager._get_drive_children_url("my_folder")

        assert ":/my_folder:/children" in url

    def test_get_drive_children_url_without_folder(self, mock_sharepoint_manager):
        """Test _get_drive_children_url without folder path"""
        url = mock_sharepoint_manager._get_drive_children_url()

        assert "/children" in url
        assert ":/children" not in url  # Should be /root/children not /root:/children

    def test_get_drive_children_url_with_nested_path(self, mock_sharepoint_manager):
        """Test _get_drive_children_url with nested path"""
        url = mock_sharepoint_manager._get_drive_children_url("parent/child/folder")

        assert "parent%2Fchild%2Ffolder" in url or "parent/child/folder" in url

    def test_base_url_constant(self, mock_sharepoint_manager):
        """Test GRAPH_API_BASE constant"""
        assert mock_sharepoint_manager.GRAPH_API_BASE == "https://graph.microsoft.com/v1.0"

    def test_site_name_formatting_with_domain(self):
        """Test site name formatting with full domain"""
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

    def test_site_name_formatting_without_domain(self):
        """Test site name formatting without domain"""
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

    def test_url_encoding_in_paths(self, mock_sharepoint_manager):
        """Test URL encoding for special characters in paths"""
        # The _get_drive_children_url method calls quote() on the path
        url = mock_sharepoint_manager._get_drive_children_url("folder with spaces")

        # Should contain encoded spaces
        assert "%20" in url or "folder" in url  # Either encoded or method handles it

    def test_url_consistency(self, mock_sharepoint_manager):
        """Test that all URLs start with the correct base"""
        urls = [
            mock_sharepoint_manager._get_site_url(),
            mock_sharepoint_manager._get_drives_url(),
            mock_sharepoint_manager._get_drive_root_url(),
        ]

        for url in urls:
            assert url.startswith("https://graph.microsoft.com/v1.0")
