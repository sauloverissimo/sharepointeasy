# SharePointEasy

✅ Easy SharePoint file operations using Microsoft Graph API. 🤙🏼

[![PyPI version](https://badge.fury.io/py/sharepointeasy.svg)](https://badge.fury.io/py/sharepointeasy)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

- **Download & Upload** - Single files or entire folders
- **File Management** - Create folders, move, copy, delete
- **Batch Operations** - Process multiple files efficiently
- **Async Support** - Full async/await API for parallel operations
- **Progress Callbacks** - Track progress of transfers
- **CLI Tool** - Command-line interface for quick operations
- **Retry & Resilience** - Automatic retry with exponential backoff
- **Version History** - List and download file versions
- **Sharing** - Create share links with permissions

## Installation

```bash
pip install sharepointeasy
```

With progress bar support (rich):
```bash
pip install sharepointeasy[rich]
```

## Quick Start

### 1. Configure credentials

Set environment variables:

```bash
export MICROSOFT_CLIENT_ID="your-client-id"
export MICROSOFT_CLIENT_SECRET="your-client-secret"
export MICROSOFT_TENANT_ID="your-tenant-id"
```

### 2. Download a file

```python
from sharepointeasy import SharePointClient

client = SharePointClient()

client.download_file(
    hostname="contoso.sharepoint.com",
    site_path="sites/MySite",
    file_path="Documents/report.xlsx",
    destination="./report.xlsx",
)
```

### 3. Upload a file

```python
client.upload_file(
    hostname="contoso.sharepoint.com",
    site_path="sites/MySite",
    file_path="Documents/new-report.xlsx",
    source="./local-report.xlsx",
)
```

## CLI Usage

After installation, use the `sharepointeasy` command:

```bash
# List sites
sharepointeasy list-sites

# List files
sharepointeasy -H contoso.sharepoint.com -S sites/MySite list Documents/

# Download file
sharepointeasy -H contoso.sharepoint.com -S sites/MySite download Documents/report.xlsx

# Upload file
sharepointeasy -H contoso.sharepoint.com -S sites/MySite upload ./report.xlsx -d Documents/

# Download entire folder
sharepointeasy -H contoso.sharepoint.com -S sites/MySite download-folder Documents/Reports

# Create folder
sharepointeasy -H contoso.sharepoint.com -S sites/MySite mkdir Documents/NewFolder -p

# Delete file
sharepointeasy -H contoso.sharepoint.com -S sites/MySite delete Documents/old-file.xlsx

# Create share link
sharepointeasy -H contoso.sharepoint.com -S sites/MySite share Documents/report.xlsx --type view
```

## Async Usage

For parallel operations, use the async client:

```python
import asyncio
from sharepointeasy import AsyncSharePointClient

async def main():
    async with AsyncSharePointClient() as client:
        # Download multiple files in parallel
        site = await client.get_site("contoso.sharepoint.com", "sites/MySite")
        drive = await client.get_drive(site["id"])

        # Download entire folder (parallel)
        await client.download_batch(
            site["id"],
            drive["id"],
            "Documents/Reports",
            "./local-reports",
            max_concurrent=10,
        )

asyncio.run(main())
```

## API Reference

### SharePointClient

#### Constructor

```python
SharePointClient(
    client_id: str | None = None,
    client_secret: str | None = None,
    tenant_id: str | None = None,
    max_retries: int = 3,
    retry_delay: float = 1.0,
)
```

#### Sites

| Method | Description |
|--------|-------------|
| `list_sites()` | List all available SharePoint sites |
| `get_site(hostname, site_path)` | Get a site by hostname and path |
| `get_site_by_name(site_name)` | Search for a site by name |

#### Drives

| Method | Description |
|--------|-------------|
| `list_drives(site_id)` | List drives (document libraries) in a site |
| `get_drive(site_id, drive_name)` | Get a drive by name |

#### Files

| Method | Description |
|--------|-------------|
| `list_files(site_id, drive_id, folder_path)` | List files in a folder |
| `list_files_recursive(site_id, drive_id, folder_path)` | List all files recursively |
| `search_file(site_id, drive_id, filename)` | Search for a file by name |
| `get_file_metadata(site_id, drive_id, file_path)` | Get file metadata |

#### Download

| Method | Description |
|--------|-------------|
| `download(site_id, drive_id, file_path, destination)` | Download a file |
| `download_batch(site_id, drive_id, folder_path, destination_dir)` | Download folder recursively |
| `download_file(hostname, site_path, file_path, destination)` | Download file (simplified) |

#### Upload

| Method | Description |
|--------|-------------|
| `upload(site_id, drive_id, file_path, source)` | Upload a file (auto chunked for large files) |
| `upload_batch(site_id, drive_id, source_dir, destination_folder)` | Upload folder recursively |
| `upload_file(hostname, site_path, file_path, source)` | Upload file (simplified) |

#### File Operations

| Method | Description |
|--------|-------------|
| `create_folder(site_id, drive_id, folder_path)` | Create a folder |
| `create_folder_recursive(site_id, drive_id, folder_path)` | Create folder with parents |
| `delete(site_id, drive_id, file_path)` | Delete a file or folder |
| `move(site_id, drive_id, source_path, destination_folder)` | Move a file or folder |
| `copy(site_id, drive_id, source_path, destination_folder)` | Copy a file or folder |

#### Versions

| Method | Description |
|--------|-------------|
| `list_versions(site_id, drive_id, file_path)` | List file versions |
| `download_version(site_id, drive_id, file_path, version_id, destination)` | Download specific version |

#### Sharing

| Method | Description |
|--------|-------------|
| `create_share_link(site_id, drive_id, file_path, link_type, scope)` | Create share link |
| `list_permissions(site_id, drive_id, file_path)` | List file permissions |

#### Lists (SharePoint Lists)

| Method | Description |
|--------|-------------|
| `list_lists(site_id)` | List all lists in a site |
| `get_list(site_id, list_name)` | Get a list by name or ID |
| `get_list_columns(site_id, list_id)` | Get list columns/fields |
| `list_items(site_id, list_id, ...)` | List items with filtering |
| `list_all_items(site_id, list_id, ...)` | List all items (auto-pagination) |
| `get_item(site_id, list_id, item_id)` | Get a specific item |
| `create_item(site_id, list_id, fields)` | Create a new item |
| `update_item(site_id, list_id, item_id, fields)` | Update an item |
| `delete_item(site_id, list_id, item_id)` | Delete an item |
| `create_list(site_id, display_name, columns)` | Create a new list |
| `delete_list(site_id, list_id)` | Delete a list |
| `batch_create_items(site_id, list_id, items)` | Create multiple items |
| `batch_update_items(site_id, list_id, updates)` | Update multiple items |
| `batch_delete_items(site_id, list_id, item_ids)` | Delete multiple items |

#### Search (Global Search)

| Method | Description |
|--------|-------------|
| `search(query, entity_types, site_id, size)` | Global search in SharePoint |
| `search_files(query, file_extension, site_id, size)` | Search files only |

#### Direct Access by ID

| Method | Description |
|--------|-------------|
| `get_item_by_id(drive_id, item_id)` | Get item metadata by ID |
| `download_by_id(drive_id, item_id, destination)` | Download file by ID |
| `upload_by_id(drive_id, parent_id, filename, source)` | Upload file to folder by ID |
| `delete_by_id(drive_id, item_id)` | Delete item by ID |

#### Microsoft Teams

| Method | Description |
|--------|-------------|
| `list_teams()` | List all teams |
| `get_team(team_id)` | Get team info |
| `list_team_channels(team_id)` | List team channels |
| `get_team_drive(team_id)` | Get team's drive |
| `list_team_files(team_id, folder_path)` | List files in team |
| `list_channel_files(team_id, channel_id)` | List files in channel |
| `download_team_file(team_id, file_path, destination)` | Download from team |
| `upload_team_file(team_id, file_path, source)` | Upload to team |

### Exceptions

| Exception | Description |
|-----------|-------------|
| `SharePointError` | Base exception |
| `AuthenticationError` | Authentication failed |
| `SiteNotFoundError` | Site not found |
| `DriveNotFoundError` | Drive not found |
| `FileNotFoundError` | File not found |
| `DownloadError` | Download failed |
| `UploadError` | Upload failed |
| `DeleteError` | Delete failed |
| `FolderCreateError` | Folder creation failed |
| `MoveError` | Move/copy failed |
| `ShareError` | Share link creation failed |
| `ListError` | List operations failed |

## Progress Tracking

### Simple callback

```python
from sharepointeasy import SharePointClient, create_progress_callback

client = SharePointClient()
progress = create_progress_callback("Downloading")

client.download_file(
    hostname="contoso.sharepoint.com",
    site_path="sites/MySite",
    file_path="Documents/large-file.zip",
    destination="./large-file.zip",
    progress_callback=progress,
)
```

### Rich progress bar (optional)

```python
from sharepointeasy import SharePointClient, RICH_AVAILABLE

if RICH_AVAILABLE:
    from sharepointeasy.utils import create_rich_progress, RichProgressCallback

    client = SharePointClient()

    with create_rich_progress() as progress:
        task = progress.add_task("Downloading...", total=100)

        # Your download with callback...
```

## Azure AD App Setup

To use this library, you need to register an app in Azure AD:

1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click **"New registration"**
3. Name your app and click **"Register"**
4. Go to **"Certificates & secrets"** → **"New client secret"**
5. Copy the secret value (this is your `MICROSOFT_CLIENT_SECRET`)
6. Go to **"API permissions"** → **"Add a permission"** → **"Microsoft Graph"**
7. Add these **Application permissions**:
   - `Sites.Read.All` - Read sites
   - `Sites.ReadWrite.All` - Read and write sites (for upload, delete, etc.)
   - `Files.Read.All` - Read files
   - `Files.ReadWrite.All` - Read and write files
8. Click **"Grant admin consent"**

Your credentials:
- `MICROSOFT_CLIENT_ID` = "Application (client) ID" on the app overview page
- `MICROSOFT_TENANT_ID` = "Directory (tenant) ID" on the app overview page
- `MICROSOFT_CLIENT_SECRET` = The secret value you copied

## Examples

### List all sites

```python
from sharepointeasy import SharePointClient

client = SharePointClient()
sites = client.list_sites()

for site in sites:
    print(f"{site['displayName']}: {site['webUrl']}")
```

### Download with error handling

```python
from sharepointeasy import (
    SharePointClient,
    FileNotFoundError,
    AuthenticationError,
)

client = SharePointClient()

try:
    client.download_file(
        hostname="contoso.sharepoint.com",
        site_path="sites/MySite",
        file_path="Documents/report.xlsx",
        destination="./report.xlsx",
    )
    print("Download completed!")
except FileNotFoundError:
    print("File not found")
except AuthenticationError:
    print("Authentication failed - check your credentials")
```

### Upload large file with progress

```python
from sharepointeasy import SharePointClient, create_progress_callback

client = SharePointClient()

# Files larger than 4MB are automatically uploaded in chunks
client.upload_file(
    hostname="contoso.sharepoint.com",
    site_path="sites/MySite",
    file_path="Backups/database.sql",
    source="./database.sql",
    progress_callback=create_progress_callback("Uploading"),
)
```

### Sync entire folder

```python
from sharepointeasy import SharePointClient

client = SharePointClient()

site = client.get_site("contoso.sharepoint.com", "sites/MySite")
drive = client.get_drive(site["id"])

# Download all files from SharePoint folder
client.download_batch(
    site["id"],
    drive["id"],
    "Documents/Project",
    "./local-project",
)

# Upload all files to SharePoint
client.upload_batch(
    site["id"],
    drive["id"],
    "./local-project",
    "Documents/Project-Backup",
)
```

### Working with SharePoint Lists

```python
from sharepointeasy import SharePointClient

client = SharePointClient()

site = client.get_site("contoso.sharepoint.com", "sites/MySite")
lst = client.get_list(site["id"], "Tasks")

# List items with filter
items = client.list_items(
    site["id"],
    lst["id"],
    filter_query="fields/Status eq 'Active'",
)

# Create a new item
client.create_item(site["id"], lst["id"], {
    "Title": "New Task",
    "Status": "Pending",
    "Priority": "High",
})

# Update an item
client.update_item(site["id"], lst["id"], "123", {
    "Status": "Completed",
})

# Batch operations
client.batch_create_items(site["id"], lst["id"], [
    {"Title": "Task 1", "Status": "New"},
    {"Title": "Task 2", "Status": "New"},
    {"Title": "Task 3", "Status": "New"},
])
```

### Async parallel downloads

```python
import asyncio
from sharepointeasy import AsyncSharePointClient

async def download_reports():
    async with AsyncSharePointClient() as client:
        site = await client.get_site("contoso.sharepoint.com", "sites/MySite")
        drive = await client.get_drive(site["id"])

        # Download 10 files in parallel
        files = await client.download_batch(
            site["id"],
            drive["id"],
            "Documents/Reports/2024",
            "./reports",
            max_concurrent=10,
        )

        print(f"Downloaded {len(files)} files")

asyncio.run(download_reports())
```

## License

MIT
