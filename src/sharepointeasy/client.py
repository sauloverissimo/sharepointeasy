"""Cliente SharePoint usando Microsoft Graph API."""

import os
import time
from pathlib import Path
from typing import BinaryIO, Callable

import httpx
from msal import ConfidentialClientApplication

from .exceptions import (
    AuthenticationError,
    DeleteError,
    DownloadError,
    DriveNotFoundError,
    FileNotFoundError,
    FolderCreateError,
    ListError,
    MoveError,
    ShareError,
    SiteNotFoundError,
    UploadError,
)


class SharePointClient:
    """Cliente para acessar SharePoint via Microsoft Graph API."""

    GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
    SCOPES = ["https://graph.microsoft.com/.default"]

    # Tamanho máximo para upload simples (4MB)
    SIMPLE_UPLOAD_MAX_SIZE = 4 * 1024 * 1024
    # Tamanho do chunk para upload de arquivos grandes (10MB)
    UPLOAD_CHUNK_SIZE = 10 * 1024 * 1024

    def __init__(
        self,
        client_id: str | None = None,
        client_secret: str | None = None,
        tenant_id: str | None = None,
        max_retries: int = 3,
        retry_delay: float = 1.0,
    ):
        """
        Inicializa o cliente SharePoint.

        Args:
            client_id: ID do aplicativo Azure AD (ou env MICROSOFT_CLIENT_ID)
            client_secret: Secret do aplicativo (ou env MICROSOFT_CLIENT_SECRET)
            tenant_id: ID do tenant Azure AD (ou env MICROSOFT_TENANT_ID)
            max_retries: Número máximo de tentativas em caso de falha
            retry_delay: Delay inicial entre tentativas (exponential backoff)

        Raises:
            AuthenticationError: Se as credenciais não estiverem configuradas
        """
        self.client_id = client_id or os.getenv("MICROSOFT_CLIENT_ID")
        self.client_secret = client_secret or os.getenv("MICROSOFT_CLIENT_SECRET")
        self.tenant_id = tenant_id or os.getenv("MICROSOFT_TENANT_ID")
        self.max_retries = max_retries
        self.retry_delay = retry_delay

        if not all([self.client_id, self.client_secret, self.tenant_id]):
            raise AuthenticationError(
                "Credenciais não configuradas. Defina MICROSOFT_CLIENT_ID, "
                "MICROSOFT_CLIENT_SECRET e MICROSOFT_TENANT_ID como variáveis de ambiente "
                "ou passe como parâmetros."
            )

        self._access_token: str | None = None
        self._token_expires_at: float = 0
        self._app: ConfidentialClientApplication | None = None

    def _get_app(self) -> ConfidentialClientApplication:
        """Retorna instância do MSAL app."""
        if self._app is None:
            authority = f"https://login.microsoftonline.com/{self.tenant_id}"
            self._app = ConfidentialClientApplication(
                client_id=self.client_id,
                client_credential=self.client_secret,
                authority=authority,
            )
        return self._app

    def _get_token(self) -> str:
        """Obtém token de acesso para Microsoft Graph com cache."""
        # Verifica se o token ainda é válido (com margem de 5 minutos)
        if self._access_token and time.time() < self._token_expires_at - 300:
            return self._access_token

        app = self._get_app()
        result = app.acquire_token_for_client(scopes=self.SCOPES)

        if "access_token" not in result:
            error = result.get("error_description", result.get("error", "Erro desconhecido"))
            raise AuthenticationError(f"Falha ao obter token: {error}")

        self._access_token = result["access_token"]
        # Token expira em 1 hora por padrão
        self._token_expires_at = time.time() + result.get("expires_in", 3600)
        return self._access_token

    def _get_headers(self) -> dict[str, str]:
        """Retorna headers para requisições à API."""
        return {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json",
        }

    def _request_with_retry(
        self,
        method: str,
        url: str,
        **kwargs,
    ) -> httpx.Response:
        """Faz requisição HTTP com retry automático."""
        last_exception = None

        for attempt in range(self.max_retries):
            try:
                with httpx.Client(timeout=60.0) as client:
                    response = client.request(
                        method,
                        url,
                        headers=self._get_headers(),
                        **kwargs,
                    )
                    # Retry em erros 429 (rate limit) e 5xx
                    if response.status_code == 429 or response.status_code >= 500:
                        retry_after = int(response.headers.get("Retry-After", self.retry_delay))
                        time.sleep(retry_after * (2 ** attempt))
                        continue
                    return response
            except httpx.HTTPError as e:
                last_exception = e
                if attempt < self.max_retries - 1:
                    time.sleep(self.retry_delay * (2 ** attempt))
                    continue
                raise

        if last_exception:
            raise last_exception
        raise RuntimeError("Falha após todas as tentativas")

    def _request(
        self,
        method: str,
        url: str,
        **kwargs,
    ) -> httpx.Response:
        """Faz requisição HTTP com headers de autenticação."""
        return self._request_with_retry(method, url, **kwargs)

    # =========================================================================
    # Sites
    # =========================================================================

    def list_sites(self) -> list[dict]:
        """
        Lista sites SharePoint disponíveis.

        Returns:
            Lista de sites com metadados
        """
        response = self._request("GET", f"{self.GRAPH_BASE_URL}/sites?search=*")
        response.raise_for_status()
        return response.json().get("value", [])

    def get_site(self, hostname: str, site_path: str) -> dict:
        """
        Obtém um site pelo hostname e path.

        Args:
            hostname: Hostname do SharePoint (ex: "contoso.sharepoint.com")
            site_path: Path do site (ex: "sites/MySite")

        Returns:
            Metadados do site

        Raises:
            SiteNotFoundError: Se o site não for encontrado
        """
        url = f"{self.GRAPH_BASE_URL}/sites/{hostname}:/{site_path}"
        response = self._request("GET", url)

        if response.status_code == 404:
            raise SiteNotFoundError(f"Site não encontrado: {hostname}/{site_path}")

        response.raise_for_status()
        return response.json()

    def get_site_by_name(self, site_name: str) -> dict | None:
        """
        Busca um site pelo nome.

        Args:
            site_name: Nome do site SharePoint

        Returns:
            Metadados do site ou None se não encontrado
        """
        sites = self.list_sites()
        for site in sites:
            if site_name.lower() in site.get("displayName", "").lower():
                return site
            if site_name.lower() in site.get("name", "").lower():
                return site
        return None

    # =========================================================================
    # Drives
    # =========================================================================

    def list_drives(self, site_id: str) -> list[dict]:
        """
        Lista drives (bibliotecas de documentos) de um site.

        Args:
            site_id: ID do site

        Returns:
            Lista de drives com metadados
        """
        response = self._request("GET", f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives")
        response.raise_for_status()
        return response.json().get("value", [])

    def get_drive(self, site_id: str, drive_name: str = "Documents") -> dict:
        """
        Obtém um drive pelo nome.

        Args:
            site_id: ID do site
            drive_name: Nome do drive (padrão: "Documents")

        Returns:
            Metadados do drive

        Raises:
            DriveNotFoundError: Se o drive não for encontrado
        """
        drives = self.list_drives(site_id)

        for drive in drives:
            name = drive.get("name", "").lower()
            if drive_name.lower() in name:
                return drive

        # Se não encontrou pelo nome, retorna o primeiro
        if drives:
            return drives[0]

        raise DriveNotFoundError(f"Nenhum drive encontrado no site {site_id}")

    # =========================================================================
    # Arquivos - Listagem e Busca
    # =========================================================================

    def list_files(
        self,
        site_id: str,
        drive_id: str,
        folder_path: str = "",
    ) -> list[dict]:
        """
        Lista arquivos em uma pasta.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            folder_path: Caminho da pasta (ex: "Documents/Reports")

        Returns:
            Lista de arquivos com metadados
        """
        if folder_path:
            url = (
                f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}"
                f"/root:/{folder_path}:/children"
            )
        else:
            url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root/children"

        response = self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    def list_files_recursive(
        self,
        site_id: str,
        drive_id: str,
        folder_path: str = "",
    ) -> list[dict]:
        """
        Lista todos os arquivos recursivamente em uma pasta.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            folder_path: Caminho da pasta

        Returns:
            Lista de todos os arquivos com metadados
        """
        all_files = []
        items = self.list_files(site_id, drive_id, folder_path)

        for item in items:
            if "folder" in item:
                # É uma pasta, listar recursivamente
                subfolder_path = f"{folder_path}/{item['name']}" if folder_path else item["name"]
                all_files.extend(self.list_files_recursive(site_id, drive_id, subfolder_path))
            else:
                # É um arquivo
                full_path = f"{folder_path}/{item['name']}" if folder_path else item["name"]
                item["_full_path"] = full_path
                all_files.append(item)

        return all_files

    def search_file(
        self,
        site_id: str,
        drive_id: str,
        filename: str,
    ) -> dict | None:
        """
        Busca um arquivo pelo nome.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            filename: Nome do arquivo

        Returns:
            Metadados do arquivo ou None se não encontrado
        """
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root/search(q='{filename}')"
        response = self._request("GET", url)
        response.raise_for_status()

        items = response.json().get("value", [])
        for item in items:
            if item.get("name", "").lower() == filename.lower():
                return item

        return None

    def get_file_metadata(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
    ) -> dict:
        """
        Obtém metadados de um arquivo.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            file_path: Caminho do arquivo

        Returns:
            Metadados do arquivo

        Raises:
            FileNotFoundError: Se o arquivo não existir
        """
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{file_path}"
        response = self._request("GET", url)

        if response.status_code == 404:
            raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

        response.raise_for_status()
        return response.json()

    # =========================================================================
    # Arquivos - Download
    # =========================================================================

    def download(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
        destination: str | Path,
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> Path:
        """
        Baixa um arquivo do SharePoint.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            file_path: Caminho do arquivo no SharePoint
            destination: Caminho local de destino
            progress_callback: Callback para progresso (bytes_downloaded, total_bytes)

        Returns:
            Path do arquivo baixado

        Raises:
            FileNotFoundError: Se o arquivo não existir
            DownloadError: Se houver erro no download
        """
        destination = Path(destination)
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{file_path}:/content"

        try:
            with httpx.Client(follow_redirects=True, timeout=120.0) as client:
                with client.stream("GET", url, headers=self._get_headers()) as response:
                    if response.status_code == 404:
                        raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

                    response.raise_for_status()

                    destination.parent.mkdir(parents=True, exist_ok=True)

                    total_size = int(response.headers.get("content-length", 0))
                    downloaded = 0

                    with open(destination, "wb") as f:
                        for chunk in response.iter_bytes(chunk_size=8192):
                            f.write(chunk)
                            downloaded += len(chunk)
                            if progress_callback and total_size:
                                progress_callback(downloaded, total_size)

        except httpx.HTTPError as e:
            raise DownloadError(f"Erro ao baixar arquivo: {e}")

        return destination

    def download_batch(
        self,
        site_id: str,
        drive_id: str,
        folder_path: str,
        destination_dir: str | Path,
        progress_callback: Callable[[str, int, int], None] | None = None,
    ) -> list[Path]:
        """
        Baixa todos os arquivos de uma pasta recursivamente.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            folder_path: Caminho da pasta no SharePoint
            destination_dir: Diretório local de destino
            progress_callback: Callback para progresso (filename, current, total)

        Returns:
            Lista de paths dos arquivos baixados
        """
        destination_dir = Path(destination_dir)
        files = self.list_files_recursive(site_id, drive_id, folder_path)
        downloaded = []

        for i, file in enumerate(files):
            file_path = file["_full_path"]
            local_path = destination_dir / file_path

            if progress_callback:
                progress_callback(file["name"], i + 1, len(files))

            self.download(site_id, drive_id, file_path, local_path)
            downloaded.append(local_path)

        return downloaded

    # =========================================================================
    # Arquivos - Upload
    # =========================================================================

    def upload(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
        source: str | Path | BinaryIO,
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> dict:
        """
        Faz upload de um arquivo para o SharePoint.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            file_path: Caminho de destino no SharePoint
            source: Arquivo local (path ou file object)
            progress_callback: Callback para progresso (bytes_uploaded, total_bytes)

        Returns:
            Metadados do arquivo criado

        Raises:
            UploadError: Se houver erro no upload
        """
        # Determinar tamanho e conteúdo
        if isinstance(source, (str, Path)):
            source = Path(source)
            file_size = source.stat().st_size
            content = source.read_bytes()
        else:
            content = source.read()
            file_size = len(content)

        try:
            if file_size <= self.SIMPLE_UPLOAD_MAX_SIZE:
                # Upload simples para arquivos pequenos
                return self._upload_simple(site_id, drive_id, file_path, content)
            else:
                # Upload em sessão para arquivos grandes
                return self._upload_large(
                    site_id, drive_id, file_path, content, progress_callback
                )
        except httpx.HTTPError as e:
            raise UploadError(f"Erro ao fazer upload: {e}")

    def _upload_simple(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
        content: bytes,
    ) -> dict:
        """Upload simples para arquivos até 4MB."""
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{file_path}:/content"

        headers = self._get_headers()
        headers["Content-Type"] = "application/octet-stream"

        with httpx.Client(timeout=120.0) as client:
            response = client.put(url, headers=headers, content=content)
            response.raise_for_status()
            return response.json()

    def _upload_large(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
        content: bytes,
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> dict:
        """Upload em sessão para arquivos grandes."""
        # Criar sessão de upload
        url = (
            f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}"
            f"/root:/{file_path}:/createUploadSession"
        )

        response = self._request("POST", url, json={
            "item": {
                "@microsoft.graph.conflictBehavior": "replace"
            }
        })
        response.raise_for_status()
        upload_url = response.json()["uploadUrl"]

        # Upload em chunks
        total_size = len(content)
        uploaded = 0

        with httpx.Client(timeout=120.0) as client:
            while uploaded < total_size:
                chunk_start = uploaded
                chunk_end = min(uploaded + self.UPLOAD_CHUNK_SIZE, total_size)
                chunk = content[chunk_start:chunk_end]

                headers = {
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {chunk_start}-{chunk_end - 1}/{total_size}",
                }

                response = client.put(upload_url, headers=headers, content=chunk)
                response.raise_for_status()

                uploaded = chunk_end
                if progress_callback:
                    progress_callback(uploaded, total_size)

            return response.json()

    def upload_batch(
        self,
        site_id: str,
        drive_id: str,
        source_dir: str | Path,
        destination_folder: str = "",
        progress_callback: Callable[[str, int, int], None] | None = None,
    ) -> list[dict]:
        """
        Faz upload de todos os arquivos de um diretório.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            source_dir: Diretório local
            destination_folder: Pasta de destino no SharePoint
            progress_callback: Callback para progresso (filename, current, total)

        Returns:
            Lista de metadados dos arquivos criados
        """
        source_dir = Path(source_dir)
        files = list(source_dir.rglob("*"))
        files = [f for f in files if f.is_file()]

        uploaded = []

        for i, local_file in enumerate(files):
            relative_path = local_file.relative_to(source_dir)
            if destination_folder:
                remote_path = f"{destination_folder}/{relative_path}"
            else:
                remote_path = str(relative_path)

            if progress_callback:
                progress_callback(local_file.name, i + 1, len(files))

            result = self.upload(site_id, drive_id, remote_path, local_file)
            uploaded.append(result)

        return uploaded

    # =========================================================================
    # Arquivos - Criar Pasta
    # =========================================================================

    def create_folder(
        self,
        site_id: str,
        drive_id: str,
        folder_path: str,
    ) -> dict:
        """
        Cria uma pasta no SharePoint.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            folder_path: Caminho da pasta a criar (ex: "Reports/2024")

        Returns:
            Metadados da pasta criada

        Raises:
            FolderCreateError: Se houver erro na criação
        """
        # Separar nome da pasta do caminho pai
        parts = folder_path.strip("/").split("/")
        folder_name = parts[-1]
        parent_path = "/".join(parts[:-1]) if len(parts) > 1 else ""

        if parent_path:
            url = (
                f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}"
                f"/root:/{parent_path}:/children"
            )
        else:
            url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root/children"

        try:
            response = self._request("POST", url, json={
                "name": folder_name,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "fail"
            })

            if response.status_code == 409:
                # Pasta já existe, retornar metadados
                return self.get_file_metadata(site_id, drive_id, folder_path)

            response.raise_for_status()
            return response.json()
        except httpx.HTTPError as e:
            raise FolderCreateError(f"Erro ao criar pasta: {e}")

    def create_folder_recursive(
        self,
        site_id: str,
        drive_id: str,
        folder_path: str,
    ) -> dict:
        """
        Cria uma pasta e todas as pastas pai necessárias.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            folder_path: Caminho completo da pasta

        Returns:
            Metadados da pasta final
        """
        parts = folder_path.strip("/").split("/")
        current_path = ""
        result = None

        for part in parts:
            current_path = f"{current_path}/{part}" if current_path else part
            try:
                result = self.create_folder(site_id, drive_id, current_path)
            except FolderCreateError:
                # Pasta já existe, continuar
                pass

        return result or self.get_file_metadata(site_id, drive_id, folder_path)

    # =========================================================================
    # Arquivos - Deletar
    # =========================================================================

    def delete(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
    ) -> bool:
        """
        Deleta um arquivo ou pasta do SharePoint.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            file_path: Caminho do arquivo/pasta

        Returns:
            True se deletado com sucesso

        Raises:
            FileNotFoundError: Se o arquivo não existir
            DeleteError: Se houver erro na deleção
        """
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{file_path}"

        try:
            response = self._request("DELETE", url)

            if response.status_code == 404:
                raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

            if response.status_code == 204:
                return True

            response.raise_for_status()
            return True
        except httpx.HTTPError as e:
            raise DeleteError(f"Erro ao deletar: {e}")

    # =========================================================================
    # Arquivos - Mover e Copiar
    # =========================================================================

    def move(
        self,
        site_id: str,
        drive_id: str,
        source_path: str,
        destination_folder: str,
        new_name: str | None = None,
    ) -> dict:
        """
        Move um arquivo ou pasta para outra localização.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            source_path: Caminho do arquivo/pasta origem
            destination_folder: Pasta de destino
            new_name: Novo nome (opcional)

        Returns:
            Metadados do item movido

        Raises:
            MoveError: Se houver erro ao mover
        """
        # Obter ID do item
        item = self.get_file_metadata(site_id, drive_id, source_path)
        item_id = item["id"]

        # Obter ID da pasta destino
        if destination_folder:
            dest = self.get_file_metadata(site_id, drive_id, destination_folder)
            parent_ref = {"id": dest["id"]}
        else:
            parent_ref = {"path": f"/drives/{drive_id}/root"}

        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/items/{item_id}"

        body = {"parentReference": parent_ref}
        if new_name:
            body["name"] = new_name

        try:
            response = self._request("PATCH", url, json=body)
            response.raise_for_status()
            return response.json()
        except httpx.HTTPError as e:
            raise MoveError(f"Erro ao mover: {e}")

    def copy(
        self,
        site_id: str,
        drive_id: str,
        source_path: str,
        destination_folder: str,
        new_name: str | None = None,
    ) -> str:
        """
        Copia um arquivo ou pasta para outra localização.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            source_path: Caminho do arquivo/pasta origem
            destination_folder: Pasta de destino
            new_name: Novo nome (opcional)

        Returns:
            URL para monitorar o progresso da cópia

        Raises:
            MoveError: Se houver erro ao copiar
        """
        # Obter ID do item
        item = self.get_file_metadata(site_id, drive_id, source_path)
        item_id = item["id"]

        # Obter ID da pasta destino
        dest = self.get_file_metadata(site_id, drive_id, destination_folder)

        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/items/{item_id}/copy"

        body = {
            "parentReference": {"driveId": drive_id, "id": dest["id"]},
        }
        if new_name:
            body["name"] = new_name
        else:
            body["name"] = item["name"]

        try:
            response = self._request("POST", url, json=body)

            if response.status_code == 202:
                # Cópia assíncrona, retornar URL de monitoramento
                return response.headers.get("Location", "")

            response.raise_for_status()
            return ""
        except httpx.HTTPError as e:
            raise MoveError(f"Erro ao copiar: {e}")

    # =========================================================================
    # Versões
    # =========================================================================

    def list_versions(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
    ) -> list[dict]:
        """
        Lista versões de um arquivo.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            file_path: Caminho do arquivo

        Returns:
            Lista de versões com metadados
        """
        item = self.get_file_metadata(site_id, drive_id, file_path)
        item_id = item["id"]

        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/items/{item_id}/versions"
        response = self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    def download_version(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
        version_id: str,
        destination: str | Path,
    ) -> Path:
        """
        Baixa uma versão específica de um arquivo.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            file_path: Caminho do arquivo
            version_id: ID da versão
            destination: Caminho local de destino

        Returns:
            Path do arquivo baixado
        """
        destination = Path(destination)
        item = self.get_file_metadata(site_id, drive_id, file_path)
        item_id = item["id"]

        url = (
            f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}"
            f"/items/{item_id}/versions/{version_id}/content"
        )

        with httpx.Client(follow_redirects=True, timeout=120.0) as client:
            response = client.get(url, headers=self._get_headers())
            response.raise_for_status()

            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(response.content)

        return destination

    # =========================================================================
    # Compartilhamento
    # =========================================================================

    def create_share_link(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
        link_type: str = "view",
        scope: str = "anonymous",
        expiration: str | None = None,
    ) -> dict:
        """
        Cria um link de compartilhamento para um arquivo.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            file_path: Caminho do arquivo
            link_type: Tipo de permissão ("view", "edit", "embed")
            scope: Escopo do link ("anonymous", "organization")
            expiration: Data de expiração ISO 8601 (opcional)

        Returns:
            Metadados do link criado

        Raises:
            ShareError: Se houver erro ao criar link
        """
        item = self.get_file_metadata(site_id, drive_id, file_path)
        item_id = item["id"]

        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/items/{item_id}/createLink"

        body = {
            "type": link_type,
            "scope": scope,
        }
        if expiration:
            body["expirationDateTime"] = expiration

        try:
            response = self._request("POST", url, json=body)
            response.raise_for_status()
            return response.json()
        except httpx.HTTPError as e:
            raise ShareError(f"Erro ao criar link de compartilhamento: {e}")

    def list_permissions(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
    ) -> list[dict]:
        """
        Lista permissões de um arquivo.

        Args:
            site_id: ID do site
            drive_id: ID do drive
            file_path: Caminho do arquivo

        Returns:
            Lista de permissões
        """
        item = self.get_file_metadata(site_id, drive_id, file_path)
        item_id = item["id"]

        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/items/{item_id}/permissions"
        response = self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    # =========================================================================
    # Listas (SharePoint Lists)
    # =========================================================================

    def list_lists(self, site_id: str) -> list[dict]:
        """
        Lista todas as listas de um site.

        Args:
            site_id: ID do site

        Returns:
            Lista de listas com metadados
        """
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists"
        response = self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    def get_list(self, site_id: str, list_name: str) -> dict:
        """
        Obtém uma lista pelo nome ou ID.

        Args:
            site_id: ID do site
            list_name: Nome ou ID da lista

        Returns:
            Metadados da lista

        Raises:
            ListError: Se a lista não for encontrada
        """
        # Tentar primeiro por ID direto
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_name}"
        response = self._request("GET", url)

        if response.status_code == 200:
            return response.json()

        # Se não encontrou, buscar por nome
        lists = self.list_lists(site_id)
        for lst in lists:
            if lst.get("displayName", "").lower() == list_name.lower():
                return lst
            if lst.get("name", "").lower() == list_name.lower():
                return lst

        raise ListError(f"Lista não encontrada: {list_name}")

    def get_list_columns(self, site_id: str, list_id: str) -> list[dict]:
        """
        Obtém as colunas (campos) de uma lista.

        Args:
            site_id: ID do site
            list_id: ID da lista

        Returns:
            Lista de colunas com metadados
        """
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}/columns"
        response = self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    def list_items(
        self,
        site_id: str,
        list_id: str,
        expand_fields: bool = True,
        filter_query: str | None = None,
        select_fields: list[str] | None = None,
        top: int | None = None,
        skip: int | None = None,
    ) -> list[dict]:
        """
        Lista itens de uma lista.

        Args:
            site_id: ID do site
            list_id: ID da lista
            expand_fields: Expandir campos (retorna valores dos campos)
            filter_query: Filtro OData (ex: "fields/Status eq 'Ativo'")
            select_fields: Campos específicos para retornar
            top: Número máximo de itens
            skip: Pular N primeiros itens

        Returns:
            Lista de itens
        """
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}/items"

        params = []
        if expand_fields:
            params.append("$expand=fields")
        if filter_query:
            params.append(f"$filter={filter_query}")
        if select_fields:
            fields_str = ",".join(select_fields)
            params.append(f"$select={fields_str}")
        if top:
            params.append(f"$top={top}")
        if skip:
            params.append(f"$skip={skip}")

        if params:
            url += "?" + "&".join(params)

        response = self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    def list_all_items(
        self,
        site_id: str,
        list_id: str,
        expand_fields: bool = True,
        filter_query: str | None = None,
    ) -> list[dict]:
        """
        Lista TODOS os itens de uma lista (com paginação automática).

        Args:
            site_id: ID do site
            list_id: ID da lista
            expand_fields: Expandir campos
            filter_query: Filtro OData

        Returns:
            Lista completa de itens
        """
        all_items = []
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}/items"

        params = []
        if expand_fields:
            params.append("$expand=fields")
        if filter_query:
            params.append(f"$filter={filter_query}")

        if params:
            url += "?" + "&".join(params)

        while url:
            response = self._request("GET", url)
            response.raise_for_status()
            data = response.json()

            all_items.extend(data.get("value", []))

            # Próxima página
            url = data.get("@odata.nextLink")

        return all_items

    def get_item(
        self,
        site_id: str,
        list_id: str,
        item_id: str,
        expand_fields: bool = True,
    ) -> dict:
        """
        Obtém um item específico de uma lista.

        Args:
            site_id: ID do site
            list_id: ID da lista
            item_id: ID do item
            expand_fields: Expandir campos

        Returns:
            Item com metadados

        Raises:
            ListError: Se o item não for encontrado
        """
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}/items/{item_id}"

        if expand_fields:
            url += "?$expand=fields"

        response = self._request("GET", url)

        if response.status_code == 404:
            raise ListError(f"Item não encontrado: {item_id}")

        response.raise_for_status()
        return response.json()

    def create_item(
        self,
        site_id: str,
        list_id: str,
        fields: dict,
    ) -> dict:
        """
        Cria um novo item em uma lista.

        Args:
            site_id: ID do site
            list_id: ID da lista
            fields: Dicionário com os campos e valores

        Returns:
            Item criado com metadados

        Raises:
            ListError: Se houver erro na criação

        Example:
            >>> client.create_item(site_id, list_id, {
            ...     "Title": "Novo Item",
            ...     "Status": "Pendente",
            ...     "Prioridade": "Alta"
            ... })
        """
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}/items"

        try:
            response = self._request("POST", url, json={"fields": fields})
            response.raise_for_status()
            return response.json()
        except httpx.HTTPError as e:
            raise ListError(f"Erro ao criar item: {e}")

    def update_item(
        self,
        site_id: str,
        list_id: str,
        item_id: str,
        fields: dict,
    ) -> dict:
        """
        Atualiza um item existente.

        Args:
            site_id: ID do site
            list_id: ID da lista
            item_id: ID do item
            fields: Campos a atualizar

        Returns:
            Item atualizado

        Raises:
            ListError: Se houver erro na atualização
        """
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"

        try:
            response = self._request("PATCH", url, json=fields)
            response.raise_for_status()
            return response.json()
        except httpx.HTTPError as e:
            raise ListError(f"Erro ao atualizar item: {e}")

    def delete_item(
        self,
        site_id: str,
        list_id: str,
        item_id: str,
    ) -> bool:
        """
        Deleta um item de uma lista.

        Args:
            site_id: ID do site
            list_id: ID da lista
            item_id: ID do item

        Returns:
            True se deletado com sucesso

        Raises:
            ListError: Se houver erro na deleção
        """
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}/items/{item_id}"

        try:
            response = self._request("DELETE", url)

            if response.status_code == 404:
                raise ListError(f"Item não encontrado: {item_id}")

            if response.status_code == 204:
                return True

            response.raise_for_status()
            return True
        except httpx.HTTPError as e:
            raise ListError(f"Erro ao deletar item: {e}")

    def create_list(
        self,
        site_id: str,
        display_name: str,
        columns: list[dict] | None = None,
        template: str = "genericList",
    ) -> dict:
        """
        Cria uma nova lista no site.

        Args:
            site_id: ID do site
            display_name: Nome da lista
            columns: Lista de definições de colunas (opcional)
            template: Template da lista (genericList, documentLibrary, etc.)

        Returns:
            Lista criada com metadados

        Example:
            >>> client.create_list(site_id, "Tarefas", columns=[
            ...     {"name": "Status", "text": {}},
            ...     {"name": "Prioridade", "choice": {"choices": ["Alta", "Média", "Baixa"]}}
            ... ])
        """
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists"

        body = {
            "displayName": display_name,
            "list": {
                "template": template,
            }
        }

        if columns:
            body["columns"] = columns

        try:
            response = self._request("POST", url, json=body)
            response.raise_for_status()
            return response.json()
        except httpx.HTTPError as e:
            raise ListError(f"Erro ao criar lista: {e}")

    def delete_list(self, site_id: str, list_id: str) -> bool:
        """
        Deleta uma lista.

        Args:
            site_id: ID do site
            list_id: ID da lista

        Returns:
            True se deletada com sucesso
        """
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}"

        try:
            response = self._request("DELETE", url)

            if response.status_code == 204:
                return True

            response.raise_for_status()
            return True
        except httpx.HTTPError as e:
            raise ListError(f"Erro ao deletar lista: {e}")

    def batch_create_items(
        self,
        site_id: str,
        list_id: str,
        items: list[dict],
    ) -> list[dict]:
        """
        Cria múltiplos itens em lote.

        Args:
            site_id: ID do site
            list_id: ID da lista
            items: Lista de dicionários com campos

        Returns:
            Lista de itens criados
        """
        created = []
        for item_fields in items:
            created.append(self.create_item(site_id, list_id, item_fields))
        return created

    def batch_update_items(
        self,
        site_id: str,
        list_id: str,
        updates: list[tuple[str, dict]],
    ) -> list[dict]:
        """
        Atualiza múltiplos itens em lote.

        Args:
            site_id: ID do site
            list_id: ID da lista
            updates: Lista de tuplas (item_id, fields)

        Returns:
            Lista de itens atualizados
        """
        updated = []
        for item_id, fields in updates:
            updated.append(self.update_item(site_id, list_id, item_id, fields))
        return updated

    def batch_delete_items(
        self,
        site_id: str,
        list_id: str,
        item_ids: list[str],
    ) -> int:
        """
        Deleta múltiplos itens em lote.

        Args:
            site_id: ID do site
            list_id: ID da lista
            item_ids: Lista de IDs dos itens

        Returns:
            Número de itens deletados
        """
        deleted = 0
        for item_id in item_ids:
            if self.delete_item(site_id, list_id, item_id):
                deleted += 1
        return deleted

    # =========================================================================
    # Search (Busca Global)
    # =========================================================================

    def search(
        self,
        query: str,
        entity_types: list[str] | None = None,
        site_id: str | None = None,
        size: int = 25,
    ) -> list[dict]:
        """
        Busca global no SharePoint.

        Args:
            query: Texto de busca
            entity_types: Tipos de entidade ("driveItem", "listItem", "site", etc.)
            site_id: Limitar busca a um site específico (opcional)
            size: Número máximo de resultados

        Returns:
            Lista de resultados da busca

        Example:
            >>> client.search("relatório vendas 2024")
            >>> client.search("*.xlsx", entity_types=["driveItem"])
        """
        url = f"{self.GRAPH_BASE_URL}/search/query"

        # Construir requisição de busca
        requests_body = {
            "queryString": query,
            "size": size,
        }

        if entity_types:
            requests_body["entityTypes"] = entity_types
        else:
            requests_body["entityTypes"] = ["driveItem", "listItem", "site"]

        # Filtrar por site se especificado
        if site_id:
            requests_body["query"] = {
                "queryString": f"{query} AND siteId:{site_id}"
            }

        body = {"requests": [requests_body]}

        response = self._request("POST", url, json=body)
        response.raise_for_status()

        # Extrair resultados
        results = []
        data = response.json()

        for search_response in data.get("value", []):
            for hit_container in search_response.get("hitsContainers", []):
                for hit in hit_container.get("hits", []):
                    results.append({
                        "id": hit.get("hitId"),
                        "rank": hit.get("rank"),
                        "summary": hit.get("summary"),
                        "resource": hit.get("resource", {}),
                    })

        return results

    def search_files(
        self,
        query: str,
        file_extension: str | None = None,
        site_id: str | None = None,
        size: int = 25,
    ) -> list[dict]:
        """
        Busca arquivos no SharePoint.

        Args:
            query: Texto de busca
            file_extension: Filtrar por extensão (ex: "xlsx", "pdf")
            site_id: Limitar a um site
            size: Número máximo de resultados

        Returns:
            Lista de arquivos encontrados
        """
        search_query = query
        if file_extension:
            search_query = f"{query} filetype:{file_extension}"

        return self.search(
            search_query,
            entity_types=["driveItem"],
            site_id=site_id,
            size=size,
        )

    # =========================================================================
    # Acesso por ID (Direct Item Access)
    # =========================================================================

    def get_item_by_id(self, drive_id: str, item_id: str) -> dict:
        """
        Obtém um arquivo ou pasta pelo ID.

        Args:
            drive_id: ID do drive
            item_id: ID do item

        Returns:
            Metadados do item

        Raises:
            FileNotFoundError: Se o item não existir
        """
        url = f"{self.GRAPH_BASE_URL}/drives/{drive_id}/items/{item_id}"
        response = self._request("GET", url)

        if response.status_code == 404:
            raise FileNotFoundError(f"Item não encontrado: {item_id}")

        response.raise_for_status()
        return response.json()

    def download_by_id(
        self,
        drive_id: str,
        item_id: str,
        destination: str | Path,
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> Path:
        """
        Baixa um arquivo pelo ID.

        Args:
            drive_id: ID do drive
            item_id: ID do arquivo
            destination: Caminho local de destino
            progress_callback: Callback para progresso

        Returns:
            Path do arquivo baixado
        """
        destination = Path(destination)
        url = f"{self.GRAPH_BASE_URL}/drives/{drive_id}/items/{item_id}/content"

        try:
            with httpx.Client(follow_redirects=True, timeout=120.0) as client:
                with client.stream("GET", url, headers=self._get_headers()) as response:
                    if response.status_code == 404:
                        raise FileNotFoundError(f"Item não encontrado: {item_id}")

                    response.raise_for_status()

                    destination.parent.mkdir(parents=True, exist_ok=True)
                    total_size = int(response.headers.get("content-length", 0))
                    downloaded = 0

                    with open(destination, "wb") as f:
                        for chunk in response.iter_bytes(chunk_size=8192):
                            f.write(chunk)
                            downloaded += len(chunk)
                            if progress_callback and total_size:
                                progress_callback(downloaded, total_size)

        except httpx.HTTPError as e:
            raise DownloadError(f"Erro ao baixar arquivo: {e}")

        return destination

    def upload_by_id(
        self,
        drive_id: str,
        parent_id: str,
        filename: str,
        source: str | Path,
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> dict:
        """
        Faz upload de um arquivo para uma pasta pelo ID.

        Args:
            drive_id: ID do drive
            parent_id: ID da pasta pai
            filename: Nome do arquivo
            source: Arquivo local
            progress_callback: Callback para progresso

        Returns:
            Metadados do arquivo criado
        """
        source = Path(source)
        file_size = source.stat().st_size
        content = source.read_bytes()

        try:
            if file_size <= self.SIMPLE_UPLOAD_MAX_SIZE:
                url = (
                    f"{self.GRAPH_BASE_URL}/drives/{drive_id}"
                    f"/items/{parent_id}:/{filename}:/content"
                )
                headers = self._get_headers()
                headers["Content-Type"] = "application/octet-stream"

                with httpx.Client(timeout=120.0) as client:
                    response = client.put(url, headers=headers, content=content)
                    response.raise_for_status()
                    return response.json()
            else:
                # Upload em sessão para arquivos grandes
                url = (
                    f"{self.GRAPH_BASE_URL}/drives/{drive_id}"
                    f"/items/{parent_id}:/{filename}:/createUploadSession"
                )
                response = self._request("POST", url, json={
                    "item": {"@microsoft.graph.conflictBehavior": "replace"}
                })
                response.raise_for_status()
                upload_url = response.json()["uploadUrl"]

                return self._upload_large_to_url(upload_url, content, progress_callback)

        except httpx.HTTPError as e:
            raise UploadError(f"Erro ao fazer upload: {e}")

    def _upload_large_to_url(
        self,
        upload_url: str,
        content: bytes,
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> dict:
        """Upload em sessão para URL pré-assinada."""
        total_size = len(content)
        uploaded = 0

        with httpx.Client(timeout=120.0) as client:
            while uploaded < total_size:
                chunk_start = uploaded
                chunk_end = min(uploaded + self.UPLOAD_CHUNK_SIZE, total_size)
                chunk = content[chunk_start:chunk_end]

                headers = {
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {chunk_start}-{chunk_end - 1}/{total_size}",
                }

                response = client.put(upload_url, headers=headers, content=chunk)
                response.raise_for_status()

                uploaded = chunk_end
                if progress_callback:
                    progress_callback(uploaded, total_size)

            return response.json()

    def delete_by_id(self, drive_id: str, item_id: str) -> bool:
        """
        Deleta um arquivo ou pasta pelo ID.

        Args:
            drive_id: ID do drive
            item_id: ID do item

        Returns:
            True se deletado com sucesso
        """
        url = f"{self.GRAPH_BASE_URL}/drives/{drive_id}/items/{item_id}"
        response = self._request("DELETE", url)

        if response.status_code == 404:
            raise FileNotFoundError(f"Item não encontrado: {item_id}")

        return response.status_code == 204

    # =========================================================================
    # Microsoft Teams
    # =========================================================================

    def list_teams(self) -> list[dict]:
        """
        Lista todos os times do Microsoft Teams.

        Returns:
            Lista de times com metadados
        """
        url = f"{self.GRAPH_BASE_URL}/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"
        response = self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    def get_team(self, team_id: str) -> dict:
        """
        Obtém informações de um time.

        Args:
            team_id: ID do time

        Returns:
            Metadados do time
        """
        url = f"{self.GRAPH_BASE_URL}/teams/{team_id}"
        response = self._request("GET", url)
        response.raise_for_status()
        return response.json()

    def list_team_channels(self, team_id: str) -> list[dict]:
        """
        Lista canais de um time.

        Args:
            team_id: ID do time

        Returns:
            Lista de canais
        """
        url = f"{self.GRAPH_BASE_URL}/teams/{team_id}/channels"
        response = self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    def get_team_drive(self, team_id: str) -> dict:
        """
        Obtém o drive (arquivos) de um time.

        Args:
            team_id: ID do time

        Returns:
            Metadados do drive do time
        """
        url = f"{self.GRAPH_BASE_URL}/groups/{team_id}/drive"
        response = self._request("GET", url)
        response.raise_for_status()
        return response.json()

    def list_team_files(
        self,
        team_id: str,
        folder_path: str = "",
    ) -> list[dict]:
        """
        Lista arquivos de um time.

        Args:
            team_id: ID do time
            folder_path: Caminho da pasta (opcional)

        Returns:
            Lista de arquivos
        """
        drive = self.get_team_drive(team_id)
        drive_id = drive["id"]

        if folder_path:
            url = f"{self.GRAPH_BASE_URL}/drives/{drive_id}/root:/{folder_path}:/children"
        else:
            url = f"{self.GRAPH_BASE_URL}/drives/{drive_id}/root/children"

        response = self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    def get_channel_files_folder(self, team_id: str, channel_id: str) -> dict:
        """
        Obtém a pasta de arquivos de um canal.

        Args:
            team_id: ID do time
            channel_id: ID do canal

        Returns:
            Metadados da pasta do canal
        """
        url = f"{self.GRAPH_BASE_URL}/teams/{team_id}/channels/{channel_id}/filesFolder"
        response = self._request("GET", url)
        response.raise_for_status()
        return response.json()

    def list_channel_files(self, team_id: str, channel_id: str) -> list[dict]:
        """
        Lista arquivos de um canal.

        Args:
            team_id: ID do time
            channel_id: ID do canal

        Returns:
            Lista de arquivos do canal
        """
        folder = self.get_channel_files_folder(team_id, channel_id)
        drive_id = folder["parentReference"]["driveId"]
        item_id = folder["id"]

        url = f"{self.GRAPH_BASE_URL}/drives/{drive_id}/items/{item_id}/children"
        response = self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    def download_team_file(
        self,
        team_id: str,
        file_path: str,
        destination: str | Path,
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> Path:
        """
        Baixa um arquivo de um time.

        Args:
            team_id: ID do time
            file_path: Caminho do arquivo no drive do time
            destination: Caminho local de destino
            progress_callback: Callback para progresso

        Returns:
            Path do arquivo baixado
        """
        drive = self.get_team_drive(team_id)
        drive_id = drive["id"]

        destination = Path(destination)
        url = f"{self.GRAPH_BASE_URL}/drives/{drive_id}/root:/{file_path}:/content"

        try:
            with httpx.Client(follow_redirects=True, timeout=120.0) as client:
                with client.stream("GET", url, headers=self._get_headers()) as response:
                    if response.status_code == 404:
                        raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

                    response.raise_for_status()

                    destination.parent.mkdir(parents=True, exist_ok=True)
                    total_size = int(response.headers.get("content-length", 0))
                    downloaded = 0

                    with open(destination, "wb") as f:
                        for chunk in response.iter_bytes(chunk_size=8192):
                            f.write(chunk)
                            downloaded += len(chunk)
                            if progress_callback and total_size:
                                progress_callback(downloaded, total_size)

        except httpx.HTTPError as e:
            raise DownloadError(f"Erro ao baixar arquivo: {e}")

        return destination

    def upload_team_file(
        self,
        team_id: str,
        file_path: str,
        source: str | Path,
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> dict:
        """
        Faz upload de um arquivo para um time.

        Args:
            team_id: ID do time
            file_path: Caminho de destino no drive do time
            source: Arquivo local
            progress_callback: Callback para progresso

        Returns:
            Metadados do arquivo criado
        """
        drive = self.get_team_drive(team_id)
        drive_id = drive["id"]

        source = Path(source)
        file_size = source.stat().st_size
        content = source.read_bytes()

        try:
            if file_size <= self.SIMPLE_UPLOAD_MAX_SIZE:
                url = f"{self.GRAPH_BASE_URL}/drives/{drive_id}/root:/{file_path}:/content"
                headers = self._get_headers()
                headers["Content-Type"] = "application/octet-stream"

                with httpx.Client(timeout=120.0) as client:
                    response = client.put(url, headers=headers, content=content)
                    response.raise_for_status()
                    return response.json()
            else:
                url = (
                    f"{self.GRAPH_BASE_URL}/drives/{drive_id}"
                    f"/root:/{file_path}:/createUploadSession"
                )
                response = self._request("POST", url, json={
                    "item": {"@microsoft.graph.conflictBehavior": "replace"}
                })
                response.raise_for_status()
                upload_url = response.json()["uploadUrl"]

                return self._upload_large_to_url(upload_url, content, progress_callback)

        except httpx.HTTPError as e:
            raise UploadError(f"Erro ao fazer upload: {e}")

    # =========================================================================
    # Métodos de conveniência
    # =========================================================================

    def download_file(
        self,
        hostname: str,
        site_path: str,
        file_path: str,
        destination: str | Path,
        drive_name: str = "Documents",
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> Path:
        """
        Baixa um arquivo usando hostname e paths (método simplificado).

        Args:
            hostname: Hostname do SharePoint (ex: "contoso.sharepoint.com")
            site_path: Path do site (ex: "sites/MySite")
            file_path: Caminho do arquivo no drive (ex: "Reports/report.xlsx")
            destination: Caminho local de destino
            drive_name: Nome do drive (padrão: "Documents")
            progress_callback: Callback para progresso

        Returns:
            Path do arquivo baixado
        """
        site = self.get_site(hostname, site_path)
        drive = self.get_drive(site["id"], drive_name)
        return self.download(site["id"], drive["id"], file_path, destination, progress_callback)

    def upload_file(
        self,
        hostname: str,
        site_path: str,
        file_path: str,
        source: str | Path,
        drive_name: str = "Documents",
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> dict:
        """
        Faz upload de um arquivo usando hostname e paths (método simplificado).

        Args:
            hostname: Hostname do SharePoint (ex: "contoso.sharepoint.com")
            site_path: Path do site (ex: "sites/MySite")
            file_path: Caminho de destino no drive (ex: "Reports/report.xlsx")
            source: Arquivo local
            drive_name: Nome do drive (padrão: "Documents")
            progress_callback: Callback para progresso

        Returns:
            Metadados do arquivo criado
        """
        site = self.get_site(hostname, site_path)
        drive = self.get_drive(site["id"], drive_name)
        return self.upload(site["id"], drive["id"], file_path, source, progress_callback)
