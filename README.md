# SharePoint Manager

A Python library for interacting with SharePoint document libraries using the Microsoft Graph API.

## Features

- 🔐 **Authentication**: MSAL-based authentication with Azure AD
- 📁 **File Operations**: Upload, download, delete, and move files
- 📂 **Folder Operations**: Download, delete, and move folders recursively
- 🔍 **Search Operations**: Search for files and folders by suffix (recursive and non-recursive)
- 🎯 **Unified API**: Single methods that auto-detect file vs folder types

## Installation

### Prerequisites

- Python 3.13+
- Microsoft Azure AD application with SharePoint permissions: `Sites.ReadWrite.All` and `
Files.ReadWrite.All`

### Dependencies

```bash
pip install msal requests python-dotenv
```

## Configuration

### 1. Azure AD App Registration

#### Steps to Register an App in Azure AD and Get Credentials:

**1. Register an App in Azure AD**

- Go to https://portal.azure.com
- Navigate to **Azure Active Directory** → **App registrations** → **New registration**
- Fill in:
    - **Name**: Something like "SharePoint Python App"
    - **Supported account types**: "Accounts in this organizational directory only"
    - **Redirect URI**: Leave blank (not needed for app-only)
- Click **Register**

**2. Get the Client ID**

- After registration, you'll see the Overview page
- Copy the **Application (client) ID** - this is your `CLIENT_ID`

**3. Create a Client Secret**

- In the same app registration, go to **Certificates & secrets** (left menu)
- Click **New client secret**
- Add a description (e.g., "Python SharePoint Access")
- Choose an expiration period (recommended: 12-24 months)
- Click **Add**
- **IMPORTANT**: Copy the **Value** immediately - this is your `CLIENT_SECRET` and you won't be able to see it again!

**4. Grant SharePoint Permissions**

- Go to **API permissions** (left menu)
- Click **Add a permission**
- Choose **SharePoint**
- Select **Application permissions** (not Delegated)
- Add these permissions:
    - `Sites.ReadWrite.All` (for read/write access to all site collections)
    - Or `Sites.FullControl.All` (for full control)
- Click **Add permissions**
- **IMPORTANT**: Click **Grant admin consent for your organization** (requires admin privileges)

**5. Get Your Tenant ID**

- Go back to **Azure Active Directory** → **Overview**
- Copy the **Tenant ID** (also called Directory ID)

### 2. Environment Variables

Create a `.env` file in your project root:

```env
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
```

## Quick Start

```python
from sharepointmanager.sharepoint import SharePointManager
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Initialize SharePoint Manager
sp_manager = SharePointManager(
    tenant_id=os.getenv('TENANT_ID'),
    client_id=os.getenv('CLIENT_ID'),
    client_secret=os.getenv('CLIENT_SECRET'),
    site_name="your-site-name"  # e.g., "contoso" for contoso.sharepoint.com
)

# Get site and drive IDs
sp_manager.get_site_id("/sites/YourSiteName")  # Optional: specify site path
sp_manager.get_drive_id("Documents")  # Document library name
```

## API Reference

### Authentication & Initialization

#### `SharePointManager(tenant_id, client_id, client_secret, site_name)`

Initialize the SharePoint Manager with authentication credentials.

**Parameters:**
- `tenant_id` (str): Azure AD tenant ID
- `client_id` (str): Application (client) ID
- `client_secret` (str): Client secret value
- `site_name` (str): SharePoint site name (without .sharepoint.com)

**Example:**
```python
sp_manager = SharePointManager(
    tenant_id="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    client_id="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
    client_secret="your-secret",
    site_name="contoso"
)
```

#### `get_site_id(site_path="")`

Retrieve the SharePoint site ID.

**Parameters:**
- `site_path` (str, optional): Site path (e.g., "/sites/YourSite")

**Returns:** Site ID (str)

#### `get_drive_id(drive_name="Documents")`

Retrieve the document library (drive) ID.

**Parameters:**
- `drive_name` (str, optional): Document library name (default: "Documents")

**Returns:** Drive ID (str)

---

### File Operations

#### `download_file(file_path, local_path=None)`

Download a file from SharePoint.

**Parameters:**
- `file_path` (str): Path to file in SharePoint (e.g., "folder/data.csv")
- `local_path` (str, optional): Local path to save file. If None, returns BytesIO object

**Returns:**
- BytesIO object (if local_path is None)
- Local file path (str) (if local_path is provided)

**Example:**
```python
# Download to memory
file_content = sp_manager.download_file("reports/data.csv")

# Download to local file
sp_manager.download_file("reports/data.csv", "/tmp/data.csv")
```

#### `upload_file(local_file_path, sharepoint_path, file_name=None)`

Upload a file to SharePoint.

**Parameters:**
- `local_file_path` (str): Path to local file
- `sharepoint_path` (str): Destination folder in SharePoint (empty string for root)
- `file_name` (str, optional): Name for file in SharePoint (defaults to local filename)

**Returns:** Response JSON from upload

**Example:**
```python
sp_manager.upload_file(
    local_file_path="/tmp/report.pdf",
    sharepoint_path="Reports/2024",
    file_name="annual_report.pdf"
)
```

#### `upload_file_from_memory(file_content, sharepoint_path, file_name)`

Upload a file from memory (bytes).

**Parameters:**
- `file_content` (bytes): File content as bytes
- `sharepoint_path` (str): Destination folder in SharePoint
- `file_name` (str): Name for the file in SharePoint

**Returns:** Response JSON from upload

**Example:**
```python
data = b"Hello, World!"
sp_manager.upload_file_from_memory(data, "Documents", "hello.txt")
```

#### `delete_file(file_path)`

Delete a file from SharePoint.

**Parameters:**
- `file_path` (str): Path to file in SharePoint

**Returns:** True if successful

**Example:**
```python
sp_manager.delete_file("old_files/temp.txt")
```

#### `move_file(file_path, destination_folder_path)`

Move a file to a different folder.

**Parameters:**
- `file_path` (str): Current path to file
- `destination_folder_path` (str): Destination folder path

**Returns:** True if successful

**Example:**
```python
sp_manager.move_file("temp/data.csv", "Archive/2024")
```

---

### Folder Operations

#### `download_folder(folder_path, local_directory=None)`

Recursively download an entire folder and its contents.

**Parameters:**
- `folder_path` (str): Path to folder in SharePoint
- `local_directory` (str, optional): Local directory to save folder (defaults to folder name)

**Returns:** Path to downloaded folder (str)

**Example:**
```python
sp_manager.download_folder("Reports/2024", "/tmp/reports")
```

#### `delete_folder(folder_path)`

Delete a folder and all its contents (permanent operation).

**Parameters:**
- `folder_path` (str): Path to folder in SharePoint

**Returns:** True if successful

**Warning:** This operation is permanent and will delete all files and subfolders.

**Example:**
```python
sp_manager.delete_folder("TempFolder")
```

#### `move_folder(folder_path, destination_parent_folder_path)`

Move a folder to a different location.

**Parameters:**
- `folder_path` (str): Current path to folder
- `destination_parent_folder_path` (str): Destination parent folder path

**Returns:** True if successful

**Example:**
```python
sp_manager.move_folder("OldProject", "Archive/Projects")
```

---

### Search Operations

#### File Search

##### `search_files_by_suffix(suffix, folder_path="")`

Search for files with a specific suffix in a folder (non-recursive).

**Parameters:**
- `suffix` (str): File suffix/extension (e.g., ".csv", ".pdf", "txt")
- `folder_path` (str, optional): Folder to search in (empty string for root)

**Returns:** List of `ItemInfo` objects

**Example:**
```python
# Search for CSV files in root
csv_files = sp_manager.search_files_by_suffix(".csv")

# Search for PDFs in specific folder
pdf_files = sp_manager.search_files_by_suffix(".pdf", "Reports/2024")

# Suffix without dot is auto-added
txt_files = sp_manager.search_files_by_suffix("txt")
```

##### `search_files_by_suffix_recursive(suffix, folder_path="")`

Recursively search for files with a specific suffix.

**Parameters:**
- `suffix` (str): File suffix/extension
- `folder_path` (str, optional): Starting folder for search

**Returns:** List of `ItemInfo` objects

**Example:**
```python
# Search all CSV files recursively from root
all_csvs = sp_manager.search_files_by_suffix_recursive(".csv")

# Search recursively from specific folder
project_docs = sp_manager.search_files_by_suffix_recursive(".docx", "Projects")
```

#### Folder Search

##### `search_folders_by_suffix(suffix, folder_path="")`

Search for folders with a specific suffix (non-recursive).

**Parameters:**
- `suffix` (str): Folder suffix (e.g., ".gdb", ".bundle")
- `folder_path` (str, optional): Folder to search in

**Returns:** List of `ItemInfo` objects

**Example:**
```python
# Search for GDB folders in root
gdb_folders = sp_manager.search_folders_by_suffix(".gdb")

# Search in specific location
backup_folders = sp_manager.search_folders_by_suffix(".backup", "Archive")
```

##### `search_folders_by_suffix_recursive(suffix, folder_path="")`

Recursively search for folders with a specific suffix.

**Parameters:**
- `suffix` (str): Folder suffix
- `folder_path` (str, optional): Starting folder for search

**Returns:** List of `ItemInfo` objects

**Example:**
```python
# Search all .gdb folders recursively
all_gdb = sp_manager.search_folders_by_suffix_recursive(".gdb")
```

---

### Unified Operations (Auto-detect File/Folder)

#### `download_item(item_path, local_path=None)`

Download a file or folder (auto-detects type).

**Parameters:**
- `item_path` (str): Path to item in SharePoint
- `local_path` (str, optional): Local path to save item

**Returns:**
- For files: BytesIO or local file path
- For folders: Path to downloaded folder

**Example:**
```python
# Downloads file or folder automatically
sp_manager.download_item("SomeItem")
```

#### `delete_item(item_path)`

Delete a file or folder (auto-detects type).

**Parameters:**
- `item_path` (str): Path to item in SharePoint

**Returns:** True if successful

**Example:**
```python
sp_manager.delete_item("OldItem")
```

#### `move_item(item_path, destination_folder_path)`

Move a file or folder (auto-detects type).

**Parameters:**
- `item_path` (str): Current path to item
- `destination_folder_path` (str): Destination folder path

**Returns:** True if successful

**Example:**
```python
sp_manager.move_item("SomeItem", "Archive")
```

---

### Data Classes

#### `ItemInfo`

Data class containing item information from SharePoint.

**Attributes:**
- `name` (str): Item name
- `path` (str): Full path to item
- `size` (int): Size in bytes
- `modified` (str): Last modified datetime (ISO format)
- `id` (str): SharePoint item ID
- `webUrl` (str): Web URL to item

**Aliases:** `FileInfo` and `FolderInfo` are aliases for `ItemInfo` (backward compatibility)

**Example:**
```python
files = sp_manager.search_files_by_suffix(".csv")
for file in files:
    print(f"Name: {file.name}")
    print(f"Path: {file.path}")
    print(f"Size: {file.size} bytes")
    print(f"Modified: {file.modified}")
    print(f"URL: {file.webUrl}")
```

---

## Complete Example

```python
from sharepointmanager.sharepoint import SharePointManager
import os
from dotenv import load_dotenv

# Load configuration
load_dotenv()

# Initialize
sp = SharePointManager(
    tenant_id=os.getenv('TENANT_ID'),
    client_id=os.getenv('CLIENT_ID'),
    client_secret=os.getenv('CLIENT_SECRET'),
    site_name="contoso"
)

# Setup site and drive
sp.get_site_id("/sites/TeamSite")
sp.get_drive_id("Documents")

# Search for files
csv_files = sp.search_files_by_suffix_recursive(".csv", "Reports")
print(f"Found {len(csv_files)} CSV files")

# Download files
for file in csv_files:
    content = sp.download_file(file.path)
    print(f"Downloaded: {file.name}")

# Upload a file
sp.upload_file("/tmp/report.pdf", "Reports/2024", "annual_report.pdf")

# Search for folders
gdb_folders = sp.search_folders_by_suffix(".gdb")
if gdb_folders:
    # Download first matching folder
    folder = gdb_folders[0]
    sp.download_folder(folder.name, "/tmp/geodatabase")

# Move items
sp.move_item("OldData.csv", "Archive/2023")

# Clean up old files
sp.delete_item("TempFolder")
```

---

## Error Handling

All methods raise exceptions on errors. Common exceptions:

- `Exception`: Generic errors (e.g., "Drive ID not set")
- `requests.exceptions.HTTPError`: HTTP errors from Graph API
  - 404: Item not found
  - 401: Authentication failed
  - 403: Permission denied

**Example:**
```python
try:
    sp.download_file("nonexistent.txt")
except requests.exceptions.HTTPError as e:
    if e.response.status_code == 404:
        print("File not found")
    else:
        print(f"Error: {e.response.status_code}")
except Exception as e:
    print(f"General error: {str(e)}")
```

---

## Testing

The library includes comprehensive pytest tests:

```bash
# Run all tests
pytest tests/

# Run specific test file
pytest tests/test_file_operations.py

# Run with verbose output
pytest tests/ -v

# Run specific test
pytest tests/test_file_operations.py::TestDownloadFile::test_download_file_to_memory
```

### Test Structure

- `tests/test_authentication_and_init.py` - Authentication and initialization
- `tests/test_file_operations.py` - File upload, download, delete, move
- `tests/test_folder_operations.py` - Folder operations and searches
- `tests/test_search_operations.py` - File search operations
- `tests/test_dataclasses.py` - ItemInfo dataclass tests
- `tests/test_url_helpers.py` - URL construction helpers
- `tests/conftest.py` - Pytest fixtures and configuration

---

## Best Practices

### 1. Always Initialize Properly

```python
# ✅ Good
sp.get_site_id()
sp.get_drive_id()
sp.download_file("file.txt")

# ❌ Bad - will raise "Drive ID not set" error
sp.download_file("file.txt")
```

### 2. Use Suffix Auto-Detection

```python
# Both work the same
sp.search_files_by_suffix(".csv")
sp.search_files_by_suffix("csv")  # Dot is added automatically
```

### 3. Use Unified Methods When Type is Unknown

```python
# Auto-detects if item is file or folder
sp.download_item("UnknownItem")
sp.delete_item("UnknownItem")
sp.move_item("UnknownItem", "Archive")
```

### 4. Handle Paths Consistently

```python
# Root folder
sp.search_files_by_suffix(".csv", "")  # Root
sp.search_files_by_suffix(".csv")      # Also root (default)

# Nested folders (no leading slash)
sp.search_files_by_suffix(".csv", "Reports/2024")
```

---

## Limitations

1. **Large File Upload**: Files larger than 4MB should use the upload session API (not currently implemented)
2. **Rate Limiting**: The Microsoft Graph API has rate limits. Consider implementing retry logic for production use
3. **Concurrent Operations**: No built-in concurrency support. Implement your own if needed
4. **Permissions**: Requires appropriate SharePoint permissions in Azure AD

---

## Troubleshooting

### Authentication Errors

**Problem:** "Authentication failed" error

**Solutions:**
- Verify Azure AD app credentials
- Check API permissions are granted and admin consented
- Ensure client secret hasn't expired

### Drive ID Not Set

**Problem:** "Drive ID not set" error

**Solution:**
```python
sp.get_drive_id()  # Call this before any file/folder operations
```

### 404 Not Found

**Problem:** Item not found errors

**Solutions:**
- Verify item path is correct (no leading slash for items)
- Check item exists in SharePoint
- Ensure you have permissions to access the item

### Site Name Issues

**Problem:** Can't connect to site

**Solutions:**
```python
# If full URL: contoso.sharepoint.com
site_name = "contoso"

# If full URL: contoso.sharepoint.com/sites/TeamSite
site_name = "contoso"
site_path = "/sites/TeamSite"
```

---

## Changelog

### Version 1.0.0 (Current)
- Initial release
- File operations (upload, download, delete, move)
- Folder operations (download, delete, move)
- Search operations (files and folders, recursive and non-recursive)
- Unified operations (auto-detect file/folder type)
- ItemInfo dataclass with backward compatible aliases
- Helper method to eliminate code duplication
- Comprehensive test suite (78 tests)

---

## License

[Add your license here]

---

## Support

For issues, questions, or contributions, please [add contact information or repository link].
