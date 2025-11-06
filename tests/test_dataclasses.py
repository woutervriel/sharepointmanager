"""
Tests for ItemInfo dataclass (FileInfo and FolderInfo are aliases)
"""
from dataclasses import asdict
from sharepointer.sharepoint import ItemInfo


class TestFileInfo:
    """Tests for FileInfo dataclass"""

    def test_file_info_creation(self):
        """Test FileInfo dataclass creation"""
        file_info = ItemInfo(
            name="test.txt",
            path="folder/test.txt",
            size=1024,
            modified="2024-01-01T00:00:00Z",
            id="file-123",
            webUrl="https://sharepoint.com/file"
        )

        assert file_info.name == "test.txt"
        assert file_info.path == "folder/test.txt"
        assert file_info.size == 1024
        assert file_info.modified == "2024-01-01T00:00:00Z"
        assert file_info.id == "file-123"
        assert file_info.webUrl == "https://sharepoint.com/file"

    def test_file_info_to_dict(self):
        """Test converting FileInfo to dictionary"""
        file_info = ItemInfo(
            name="test.txt",
            path="folder/test.txt",
            size=1024,
            modified="2024-01-01T00:00:00Z",
            id="file-123",
            webUrl="https://sharepoint.com/file"
        )

        file_dict = asdict(file_info)

        assert isinstance(file_dict, dict)
        assert file_dict["name"] == "test.txt"
        assert file_dict["size"] == 1024

    def test_file_info_with_zero_size(self):
        """Test FileInfo with zero size"""
        file_info = ItemInfo(
            name="empty.txt",
            path="folder/empty.txt",
            size=0,
            modified="2024-01-01T00:00:00Z",
            id="file-456",
            webUrl="https://sharepoint.com/empty"
        )

        assert file_info.size == 0

    def test_file_info_with_special_characters(self):
        """Test FileInfo with special characters in name"""
        file_info = ItemInfo(
            name="test_file (1) [2024].txt",
            path="folder/test_file (1) [2024].txt",
            size=1024,
            modified="2024-01-01T00:00:00Z",
            id="file-789",
            webUrl="https://sharepoint.com/file"
        )

        assert "(" in file_info.name
        assert "[" in file_info.name

