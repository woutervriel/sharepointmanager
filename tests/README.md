# SharePoint Manager Tests

This directory contains comprehensive test suite for the SharePoint Manager using pytest.

## Test Structure

```
tests/
├── __init__.py                  # Package marker
├── conftest.py                  # Shared fixtures and configuration
├── test_dataclasses.py         # Tests for FileInfo and FolderInfo dataclasses
├── test_url_helpers.py         # Tests for URL builder methods
├── test_file_operations.py     # Tests for file operations (download, upload, delete, move)
├── test_folder_operations.py   # Tests for folder operations
├── test_search_operations.py   # Tests for search operations
└── README.md                    # This file
```

## Setup

### Install Test Dependencies

```bash
pip install -r requirements-test.txt
```

Or install only pytest and dependencies:

```bash
pip install pytest pytest-cov pytest-mock
```

## Running Tests

### Run All Tests

```bash
pytest
```

### Run with Verbose Output

```bash
pytest -v
```

### Run Specific Test File

```bash
pytest tests/test_dataclasses.py
```

### Run Specific Test Class

```bash
pytest tests/test_dataclasses.py::TestFileInfo
```

### Run Specific Test

```bash
pytest tests/test_dataclasses.py::TestFileInfo::test_file_info_creation
```

### Run with Coverage Report

```bash
pytest --cov=main5 --cov-report=html
```

This generates an HTML coverage report in `htmlcov/index.html`

### Run Only Unit Tests

```bash
pytest -m unit
```

### Run Only Integration Tests

```bash
pytest -m integration
```

### Stop on First Failure

```bash
pytest -x
```

### Show Print Statements

```bash
pytest -s
```

## Test Files Overview

### test_dataclasses.py
Tests for the `FileInfo` and `FolderInfo` dataclasses:
- Creation and initialization
- Conversion to dictionaries using `asdict()`
- Equality comparison
- Handling of special characters and nested paths

**Classes:**
- `TestFileInfo` - Tests for FileInfo dataclass
- `TestFolderInfo` - Tests for FolderInfo dataclass
- `TestDataclassEquality` - Tests for equality operations

### test_url_helpers.py
Tests for all URL builder methods:
- `_get_site_url()` - Site URL building
- `_get_drives_url()` - Drives list URL
- `_get_drive_root_url()` - Drive root URL
- `_get_drive_item_url()` - Item URL with path
- `_get_drive_item_content_url()` - Content URL for uploads
- `_get_drive_children_url()` - Children URL for folder listing
- Site name formatting
- URL encoding for special characters

**Classes:**
- `TestURLHelpers` - All URL helper tests

### test_file_operations.py
Tests for file-related operations:
- `download_file()` - Download to memory and local path
- `upload_file()` - Upload from local path
- `upload_file_from_memory()` - Upload from bytes
- `delete_file()` - Delete files
- `move_file()` - Move files to different folders
- Error handling (404, missing drive_id, etc.)

**Classes:**
- `TestDownloadFile` - File download tests
- `TestUploadFile` - File upload tests
- `TestDeleteFile` - File deletion tests
- `TestMoveFile` - File move tests

### test_folder_operations.py
Tests for folder-related operations:
- `search_folders_by_suffix_recursive()` - Find folders by suffix
- `download_folder()` - Download folder with contents
- `delete_folder()` - Delete folder and contents
- `move_folder()` - Move folders to different locations
- Error handling and edge cases

**Classes:**
- `TestSearchFoldersByPrefix` - Folder search tests
- `TestDeleteFolder` - Folder deletion tests
- `TestMoveFolder` - Folder move tests
- `TestDownloadFolder` - Folder download tests

### test_search_operations.py
Tests for search operations:
- `search_files_by_suffix()` - Search files in folder
- `search_files_by_suffix_recursive()` - Search files recursively
- Multiple matches, no matches, empty suffix
- Input validation
- Correct dataclass returns

**Classes:**
- `TestSearchFilesBySuffix` - Basic file search tests
- `TestSearchFilesByRecursive` - Recursive search tests
- `TestSearchFilesBySuffixValidation` - Input validation tests

## Fixtures (conftest.py)

The `conftest.py` file provides shared fixtures for all tests:

### Mock Fixtures
- `mock_sharepoint_manager` - Mocked SharePointManager instance
- `mock_requests_get` - Mocked requests.get
- `mock_requests_post` - Mocked requests.post
- `mock_requests_patch` - Mocked requests.patch
- `mock_requests_put` - Mocked requests.put
- `mock_requests_delete` - Mocked requests.delete

### Data Fixtures
- `sample_file_info` - Sample FileInfo object
- `sample_folder_info` - Sample FolderInfo object
- `mock_response_200` - Successful (200) response mock
- `mock_response_404` - Not found (404) response mock
- `mock_response_with_download_url` - Response with download URL

## Test Coverage

Current test coverage includes:

- **Dataclasses**: 100%
  - FileInfo creation and operations
  - FolderInfo creation and operations
  - Equality and comparison

- **URL Helpers**: 100%
  - All URL building methods
  - Edge cases (empty paths, special characters)
  - URL consistency

- **File Operations**: ~90%
  - Download, upload, delete, move
  - Error handling
  - Local path and in-memory operations

- **Folder Operations**: ~85%
  - Search, download, delete, move
  - Error handling
  - Path handling

- **Search Operations**: ~90%
  - File search (single and recursive)
  - Multiple matches
  - Input validation

## Mocking Strategy

All tests use mocking to avoid actual API calls to SharePoint:

1. **Authentication**: MSAL is mocked to avoid needing real credentials
2. **HTTP Requests**: `requests` module methods are mocked
3. **File I/O**: File operations are mocked where needed
4. **Responses**: Mock response objects simulate SharePoint API responses

## Common Test Patterns

### Testing Successful Operations

```python
def test_operation_success(self, mock_sharepoint_manager, mock_requests_get):
    response = Mock()
    response.json.return_value = {"id": "123"}
    response.raise_for_status.return_value = None

    mock_requests_get.return_value = response

    result = mock_sharepoint_manager.some_method()
    assert result is not None
```

### Testing Error Handling

```python
def test_operation_error(self, mock_sharepoint_manager, mock_requests_get):
    response = Mock()
    response.status_code = 404
    response.raise_for_status.side_effect = HTTPError(response=response)

    mock_requests_get.return_value = response

    with pytest.raises(HTTPError):
        mock_sharepoint_manager.some_method()
```

### Testing Without Required IDs

```python
def test_without_drive_id(self):
    with patch('main5.msal.ConfidentialClientApplication'):
        manager = SharePointManager("tenant", "client", "secret", "site")
        manager.access_token = "token"
        manager.site_id = "site-id"
        manager.drive_id = None

        with pytest.raises(Exception, match="Drive ID not set"):
            manager.some_method()
```

## Best Practices

1. **Use Fixtures**: Always use fixtures from conftest.py for consistency
2. **Test One Thing**: Each test should test a single behavior
3. **Clear Assertions**: Make assertions specific and clear
4. **Test Error Cases**: Always test both success and failure paths
5. **Use Mocking**: Never make actual API calls in unit tests
6. **Descriptive Names**: Use clear, descriptive test names

## Troubleshooting

### Import Errors
If you get import errors, make sure:
- You're running pytest from the project root directory
- The `conftest.py` file is in the tests directory
- Python path includes the project root

### Fixture Not Found
If fixtures aren't found:
- Check that `conftest.py` is in the tests directory
- Ensure fixture names match exactly (case-sensitive)
- Use `pytest --fixtures` to see all available fixtures

### Mock Not Working
If mocks aren't working:
- Check patch paths are correct (should be `main5.module_name`)
- Ensure mocks are set up before calling the method
- Verify return values are set correctly

## CI/CD Integration

To integrate with CI/CD:

```bash
# Run tests with coverage
pytest --cov=main5 --cov-report=xml

# Run tests and exit with status code
pytest --tb=short

# Run tests in parallel (requires pytest-xdist)
pytest -n auto
```

## Contributing

When adding new features to SharePointManager:

1. Write tests first (TDD approach)
2. Ensure all tests pass: `pytest`
3. Check coverage: `pytest --cov=main5`
4. Maintain >85% coverage
5. Document test purpose in docstrings
