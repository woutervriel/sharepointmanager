import os

from dotenv import load_dotenv

from sharepointmanager.sharepoint import SharePointManager

load_dotenv()


# Configuration
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
TENANT_ID = os.getenv('TENANT_ID')

SITE_NAME = os.getenv('SITE_NAME')  # Your SharePoint tenant name (or full URL)
SITE_PATH = os.getenv('SITE_PATH')  # Site path (e.g., '/sites/yoursite' or '' for root)
DRIVE_NAME = os.getenv('DRIVE_NAME')  # Document library name (usually 'Documents')


try:
    # Initialize SharePoint manager
    print("=" * 60)
    print("Connecting to SharePoint using MSAL...")
    print("=" * 60)
    sp_manager = SharePointManager(TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_NAME)

    # Get site and drive IDs
    print("\nRetrieving site information...")
    sp_manager.get_site_id(SITE_PATH)
    sp_manager.get_drive_id(DRIVE_NAME)

    # Search for .gdb folder in SharePoint
    print("\n" + "=" * 60)
    print("Searching folder from SharePoint...")
    print("=" * 60)
    matches = sp_manager.search_folders_by_suffix(".csv", folder_path='')

    if len(matches) == 0:
        print("✗ No .csv folders found")
    else:
        to_download = matches[0].name
        print(f"✓ Found folder: {to_download}")

        # Download folder from SharePoint
        print("\n" + "=" * 60)
        print("Downloading folder from SharePoint...")
        print("=" * 60)
        downloaded_path = sp_manager.download_folder(to_download)

    # Move folder in SharePoint
    print("\n" + "=" * 60)
    print("Moving folder in SharePoint...")
    print("=" * 60)
    destination_folder = "Archive"
    sp_manager.move_item(item_path=to_download, destination_folder_path=destination_folder)

    print("\n" + "=" * 60)
    print("✓ Process completed successfully!")
    print("=" * 60)

except Exception as e:
    print("\n" + "=" * 60)
    print(f"✗ Error occurred: {str(e)}")
    print("=" * 60)
    raise


