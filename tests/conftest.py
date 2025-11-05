"""
Pytest configuration and shared fixtures for SharePoint Manager tests
"""
import pytest
from unittest.mock import Mock, patch
import sys
import os

# Add parent directory to path so we can import sharepointmanager.sharepoint
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dotenv import load_dotenv

load_dotenv()

# Configuration
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID = os.getenv('TENANT_ID')

SITE_NAME = os.getenv('SITE_NAME')  # Your SharePoint tenant name (or full URL)
SITE_PATH = os.getenv('SITE_PATH')  # Site path (e.g., '/sites/yoursite' or '' for root)
DRIVE_NAME = os.getenv('DRIVE_NAME')  # Document library name (usually 'Documents')


@pytest.fixture
def config():
    """Fixture that provides configuration data"""
    return {
        "CLIENT_ID": CLIENT_ID,
        "CLIENT_SECRET": CLIENT_SECRET,
        "TENANT_ID": TENANT_ID,
        "SITE_NAME": SITE_NAME,
        "SITE_PATH": SITE_PATH,
        "DRIVE_NAME": DRIVE_NAME
    }


@pytest.fixture
def mock_sharepoint_manager(config):
    """Fixture that provides a mocked SharePointManager instance"""
    with patch('sharepointmanager.sharepoint.msal.ConfidentialClientApplication') as mock_auth:
        from sharepointmanager.sharepoint import SharePointManager

        # Configure the mock to return a proper access token
        mock_app = Mock()
        mock_app.acquire_token_for_client.return_value = {"access_token": "test-token"}
        mock_auth.return_value = mock_app

        manager = SharePointManager(
            tenant_id=config["TENANT_ID"],
            client_id=config["CLIENT_ID"],
            client_secret=config["CLIENT_SECRET"],
            site_name=config["SITE_NAME"]
        )

        # Set the required IDs
        manager.site_id = "test-site-id"
        manager.drive_id = "test-drive-id"

        return manager


@pytest.fixture
def mock_requests_get():
    """Fixture that provides a mock for requests.get"""
    with patch('sharepointmanager.sharepoint.requests.get') as mock:
        yield mock


@pytest.fixture
def mock_requests_post():
    """Fixture that provides a mock for requests.post"""
    with patch('sharepointmanager.sharepoint.requests.post') as mock:
        yield mock


@pytest.fixture
def mock_requests_patch():
    """Fixture that provides a mock for requests.patch"""
    with patch('sharepointmanager.sharepoint.requests.patch') as mock:
        yield mock


@pytest.fixture
def mock_requests_put():
    """Fixture that provides a mock for requests.put"""
    with patch('sharepointmanager.sharepoint.requests.put') as mock:
        yield mock


@pytest.fixture
def mock_requests_delete():
    """Fixture that provides a mock for requests.delete"""
    with patch('sharepointmanager.sharepoint.requests.delete') as mock:
        yield mock


@pytest.fixture
def sample_file_info():
    """Fixture providing sample FileInfo data"""
    from sharepointmanager.sharepoint import FileInfo
    return FileInfo(
        name="test_file.txt",
        path="folder/test_file.txt",
        size=1024,
        modified="2024-01-01T00:00:00Z",
        id="file-id-123",
        webUrl="https://sharepoint.com/file"
    )


@pytest.fixture
def sample_folder_info():
    """Fixture providing sample FolderInfo data"""
    from sharepointmanager.sharepoint import FolderInfo
    return FolderInfo(
        name="test_folder",
        path="folder/test_folder",
        size=5120,
        modified="2024-01-01T00:00:00Z",
        id="folder-id-456",
        webUrl="https://sharepoint.com/folder"
    )


@pytest.fixture
def mock_response_200():
    """Fixture providing a successful response mock"""
    response = Mock()
    response.status_code = 200
    response.json.return_value = {"id": "test-id", "name": "test"}
    response.raise_for_status.return_value = None
    return response


@pytest.fixture
def mock_response_404():
    """Fixture providing a 404 response mock"""
    response = Mock()
    response.status_code = 404
    response.text = "Not Found"
    from requests.exceptions import HTTPError
    response.raise_for_status.side_effect = HTTPError(response=response)
    return response


@pytest.fixture
def mock_response_with_download_url():
    """Fixture providing a response with download URL"""
    response = Mock()
    response.status_code = 200
    response.json.return_value = {
        "id": "test-id",
        "name": "test_file.txt",
        "@microsoft.graph.downloadUrl": "https://download.sharepoint.com/file"
    }
    response.raise_for_status.return_value = None
    return response
