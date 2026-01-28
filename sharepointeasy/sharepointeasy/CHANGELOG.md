# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.4.0] - 2026-01-28

### Added
- Full async client (`AsyncSharePointClient`) with parallel operations
- Microsoft Teams integration (list teams, channels, files)
- SharePoint Lists support (CRUD operations, batch operations)
- Global search functionality
- Direct item access by ID
- Version history support (list and download versions)
- Sharing links creation with permissions
- Progress callbacks for uploads and downloads
- Rich progress bar support (optional dependency)
- CLI with 28+ commands and intuitive aliases

### Changed
- Environment variables renamed from `MICROSOFT_CLIENTE_*` to `MICROSOFT_CLIENT_*`
- CLI now uses text indicators `[DIR]`/`[FILE]` instead of emojis for compatibility
- Improved error messages with examples

### Fixed
- Long lines exceeding 100 characters (ruff compliance)
- Removed unused `BinaryIO` import

## [0.3.0] - 2025-12-15

### Added
- Batch download/upload operations
- Folder creation with recursive parent creation
- Move and copy operations
- Delete operations for files and folders

## [0.2.0] - 2025-11-20

### Added
- Basic download and upload functionality
- Site and drive listing
- File listing (recursive and non-recursive)
- File search by name

## [0.1.0] - 2025-10-01

### Added
- Initial release
- Authentication with Azure AD
- Basic SharePoint connection
