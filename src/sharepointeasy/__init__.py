"""SharePointEasy - Easy SharePoint file operations using Microsoft Graph API."""

from .async_client import AsyncSharePointClient
from .client import SharePointClient
from .exceptions import (
    AuthenticationError,
    CalendarError,
    DeleteError,
    DownloadError,
    DriveNotFoundError,
    EventNotFoundError,
    FileNotFoundError,
    FolderCreateError,
    ListError,
    MoveError,
    ShareError,
    SharePointError,
    SiteNotFoundError,
    UploadError,
)
from .utils import (
    RICH_AVAILABLE,
    create_batch_progress_callback,
    create_progress_callback,
    format_size,
)

__version__ = "1.0.0"
__all__ = [
    # Clients
    "SharePointClient",
    "AsyncSharePointClient",
    # Exceptions
    "SharePointError",
    "AuthenticationError",
    "SiteNotFoundError",
    "DriveNotFoundError",
    "FileNotFoundError",
    "DownloadError",
    "UploadError",
    "DeleteError",
    "FolderCreateError",
    "MoveError",
    "ShareError",
    "ListError",
    "CalendarError",
    "EventNotFoundError",
    # Utils
    "create_progress_callback",
    "create_batch_progress_callback",
    "format_size",
    "RICH_AVAILABLE",
]
