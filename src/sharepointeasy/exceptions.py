"""Exceções customizadas para sharepointeasy."""


class SharePointError(Exception):
    """Erro base para operações SharePoint."""

    pass


class AuthenticationError(SharePointError):
    """Erro de autenticação com Microsoft Graph API."""

    pass


class SiteNotFoundError(SharePointError):
    """Site SharePoint não encontrado."""

    pass


class DriveNotFoundError(SharePointError):
    """Drive/biblioteca de documentos não encontrado."""

    pass


class FileNotFoundError(SharePointError):
    """Arquivo não encontrado no SharePoint."""

    pass


class DownloadError(SharePointError):
    """Erro ao baixar arquivo."""

    pass


class UploadError(SharePointError):
    """Erro ao fazer upload de arquivo."""

    pass


class DeleteError(SharePointError):
    """Erro ao deletar arquivo ou pasta."""

    pass


class FolderCreateError(SharePointError):
    """Erro ao criar pasta."""

    pass


class MoveError(SharePointError):
    """Erro ao mover ou copiar arquivo."""

    pass


class ShareError(SharePointError):
    """Erro ao criar link de compartilhamento."""

    pass


class ListError(SharePointError):
    """Erro em operações com listas do SharePoint."""

    pass
