"""Cliente SharePoint assíncrono usando Microsoft Graph API."""

import asyncio
import os
import time
from pathlib import Path
from typing import Callable

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


class AsyncSharePointClient:
    """Cliente assíncrono para acessar SharePoint via Microsoft Graph API."""

    GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
    SCOPES = ["https://graph.microsoft.com/.default"]

    SIMPLE_UPLOAD_MAX_SIZE = 4 * 1024 * 1024
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
        Inicializa o cliente SharePoint assíncrono.

        Args:
            client_id: ID do aplicativo Azure AD (ou env MICROSOFT_CLIENT_ID)
            client_secret: Secret do aplicativo (ou env MICROSOFT_CLIENT_SECRET)
            tenant_id: ID do tenant Azure AD (ou env MICROSOFT_TENANT_ID)
            max_retries: Número máximo de tentativas em caso de falha
            retry_delay: Delay inicial entre tentativas (exponential backoff)
        """
        self.client_id = client_id or os.getenv("MICROSOFT_CLIENT_ID")
        self.client_secret = client_secret or os.getenv("MICROSOFT_CLIENT_SECRET")
        self.tenant_id = tenant_id or os.getenv("MICROSOFT_TENANT_ID")
        self.max_retries = max_retries
        self.retry_delay = retry_delay

        if not all([self.client_id, self.client_secret, self.tenant_id]):
            raise AuthenticationError(
                "Credenciais não configuradas. Defina MICROSOFT_CLIENT_ID, "
                "MICROSOFT_CLIENT_SECRET e MICROSOFT_TENANT_ID."
            )

        self._access_token: str | None = None
        self._token_expires_at: float = 0
        self._app: ConfidentialClientApplication | None = None
        self._client: httpx.AsyncClient | None = None

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
        """Obtém token de acesso para Microsoft Graph."""
        if self._access_token and time.time() < self._token_expires_at - 300:
            return self._access_token

        app = self._get_app()
        result = app.acquire_token_for_client(scopes=self.SCOPES)

        if "access_token" not in result:
            error = result.get("error_description", result.get("error", "Erro desconhecido"))
            raise AuthenticationError(f"Falha ao obter token: {error}")

        self._access_token = result["access_token"]
        self._token_expires_at = time.time() + result.get("expires_in", 3600)
        return self._access_token

    def _get_headers(self) -> dict[str, str]:
        """Retorna headers para requisições à API."""
        return {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json",
        }

    async def __aenter__(self):
        """Context manager entry."""
        self._client = httpx.AsyncClient(timeout=60.0)
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        if self._client:
            await self._client.aclose()
            self._client = None

    async def _request(
        self,
        method: str,
        url: str,
        **kwargs,
    ) -> httpx.Response:
        """Faz requisição HTTP assíncrona com retry."""
        if not self._client:
            raise RuntimeError("Use 'async with' para gerenciar o cliente")

        last_exception = None

        for attempt in range(self.max_retries):
            try:
                response = await self._client.request(
                    method,
                    url,
                    headers=self._get_headers(),
                    **kwargs,
                )

                if response.status_code == 429 or response.status_code >= 500:
                    retry_after = int(response.headers.get("Retry-After", self.retry_delay))
                    await asyncio.sleep(retry_after * (2 ** attempt))
                    continue

                return response

            except httpx.HTTPError as e:
                last_exception = e
                if attempt < self.max_retries - 1:
                    await asyncio.sleep(self.retry_delay * (2 ** attempt))
                    continue
                raise

        if last_exception:
            raise last_exception
        raise RuntimeError("Falha após todas as tentativas")

    # =========================================================================
    # Sites
    # =========================================================================

    async def list_sites(self) -> list[dict]:
        """Lista sites SharePoint disponíveis."""
        response = await self._request("GET", f"{self.GRAPH_BASE_URL}/sites?search=*")
        response.raise_for_status()
        return response.json().get("value", [])

    async def get_site(self, hostname: str, site_path: str) -> dict:
        """Obtém um site pelo hostname e path."""
        url = f"{self.GRAPH_BASE_URL}/sites/{hostname}:/{site_path}"
        response = await self._request("GET", url)

        if response.status_code == 404:
            raise SiteNotFoundError(f"Site não encontrado: {hostname}/{site_path}")

        response.raise_for_status()
        return response.json()

    async def get_site_by_name(self, site_name: str) -> dict | None:
        """Busca um site pelo nome."""
        sites = await self.list_sites()
        for site in sites:
            if site_name.lower() in site.get("displayName", "").lower():
                return site
            if site_name.lower() in site.get("name", "").lower():
                return site
        return None

    # =========================================================================
    # Drives
    # =========================================================================

    async def list_drives(self, site_id: str) -> list[dict]:
        """Lista drives de um site."""
        response = await self._request("GET", f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives")
        response.raise_for_status()
        return response.json().get("value", [])

    async def get_drive(self, site_id: str, drive_name: str = "Documents") -> dict:
        """Obtém um drive pelo nome."""
        drives = await self.list_drives(site_id)

        for drive in drives:
            name = drive.get("name", "").lower()
            if drive_name.lower() in name:
                return drive

        if drives:
            return drives[0]

        raise DriveNotFoundError(f"Nenhum drive encontrado no site {site_id}")

    # =========================================================================
    # Arquivos - Listagem
    # =========================================================================

    async def list_files(
        self,
        site_id: str,
        drive_id: str,
        folder_path: str = "",
    ) -> list[dict]:
        """Lista arquivos em uma pasta."""
        if folder_path:
            url = (
                f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}"
                f"/root:/{folder_path}:/children"
            )
        else:
            url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root/children"

        response = await self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    async def list_files_recursive(
        self,
        site_id: str,
        drive_id: str,
        folder_path: str = "",
    ) -> list[dict]:
        """Lista todos os arquivos recursivamente."""
        all_files = []
        items = await self.list_files(site_id, drive_id, folder_path)

        tasks = []
        for item in items:
            if "folder" in item:
                if folder_path:
                    subfolder_path = f"{folder_path}/{item['name']}"
                else:
                    subfolder_path = item["name"]
                tasks.append(self.list_files_recursive(site_id, drive_id, subfolder_path))
            else:
                if folder_path:
                    item["_full_path"] = f"{folder_path}/{item['name']}"
                else:
                    item["_full_path"] = item["name"]
                all_files.append(item)

        if tasks:
            results = await asyncio.gather(*tasks)
            for result in results:
                all_files.extend(result)

        return all_files

    async def get_file_metadata(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
    ) -> dict:
        """Obtém metadados de um arquivo."""
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{file_path}"
        response = await self._request("GET", url)

        if response.status_code == 404:
            raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

        response.raise_for_status()
        return response.json()

    # =========================================================================
    # Download
    # =========================================================================

    async def download(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
        destination: str | Path,
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> Path:
        """Baixa um arquivo do SharePoint."""
        destination = Path(destination)
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{file_path}:/content"

        try:
            async with httpx.AsyncClient(follow_redirects=True, timeout=120.0) as client:
                async with client.stream("GET", url, headers=self._get_headers()) as response:
                    if response.status_code == 404:
                        raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

                    response.raise_for_status()

                    destination.parent.mkdir(parents=True, exist_ok=True)
                    total_size = int(response.headers.get("content-length", 0))
                    downloaded = 0

                    with open(destination, "wb") as f:
                        async for chunk in response.aiter_bytes(chunk_size=8192):
                            f.write(chunk)
                            downloaded += len(chunk)
                            if progress_callback and total_size:
                                progress_callback(downloaded, total_size)

        except httpx.HTTPError as e:
            raise DownloadError(f"Erro ao baixar arquivo: {e}")

        return destination

    async def download_batch(
        self,
        site_id: str,
        drive_id: str,
        folder_path: str,
        destination_dir: str | Path,
        max_concurrent: int = 5,
        progress_callback: Callable[[str, int, int], None] | None = None,
    ) -> list[Path]:
        """Baixa múltiplos arquivos em paralelo."""
        destination_dir = Path(destination_dir)
        files = await self.list_files_recursive(site_id, drive_id, folder_path)

        semaphore = asyncio.Semaphore(max_concurrent)
        downloaded = []
        completed = 0

        async def download_one(file: dict) -> Path:
            nonlocal completed
            async with semaphore:
                file_path = file["_full_path"]
                local_path = destination_dir / file_path

                # Criar novo cliente para cada download paralelo
                async with httpx.AsyncClient(follow_redirects=True, timeout=120.0) as client:
                    url = (
                        f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}"
                        f"/root:/{file_path}:/content"
                    )
                    async with client.stream("GET", url, headers=self._get_headers()) as response:
                        response.raise_for_status()
                        local_path.parent.mkdir(parents=True, exist_ok=True)

                        with open(local_path, "wb") as f:
                            async for chunk in response.aiter_bytes(chunk_size=8192):
                                f.write(chunk)

                completed += 1
                if progress_callback:
                    progress_callback(file["name"], completed, len(files))

                return local_path

        tasks = [download_one(f) for f in files]
        downloaded = await asyncio.gather(*tasks)

        return list(downloaded)

    # =========================================================================
    # Upload
    # =========================================================================

    async def upload(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
        source: str | Path,
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> dict:
        """Faz upload de um arquivo."""
        source = Path(source)
        file_size = source.stat().st_size
        content = source.read_bytes()

        try:
            if file_size <= self.SIMPLE_UPLOAD_MAX_SIZE:
                return await self._upload_simple(site_id, drive_id, file_path, content)
            else:
                return await self._upload_large(
                    site_id, drive_id, file_path, content, progress_callback
                )
        except httpx.HTTPError as e:
            raise UploadError(f"Erro ao fazer upload: {e}")

    async def _upload_simple(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
        content: bytes,
    ) -> dict:
        """Upload simples para arquivos pequenos."""
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{file_path}:/content"

        headers = self._get_headers()
        headers["Content-Type"] = "application/octet-stream"

        async with httpx.AsyncClient(timeout=120.0) as client:
            response = await client.put(url, headers=headers, content=content)
            response.raise_for_status()
            return response.json()

    async def _upload_large(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
        content: bytes,
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> dict:
        """Upload em sessão para arquivos grandes."""
        url = (
            f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}"
            f"/root:/{file_path}:/createUploadSession"
        )

        response = await self._request("POST", url, json={
            "item": {"@microsoft.graph.conflictBehavior": "replace"}
        })
        response.raise_for_status()
        upload_url = response.json()["uploadUrl"]

        total_size = len(content)
        uploaded = 0

        async with httpx.AsyncClient(timeout=120.0) as client:
            while uploaded < total_size:
                chunk_start = uploaded
                chunk_end = min(uploaded + self.UPLOAD_CHUNK_SIZE, total_size)
                chunk = content[chunk_start:chunk_end]

                headers = {
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {chunk_start}-{chunk_end - 1}/{total_size}",
                }

                response = await client.put(upload_url, headers=headers, content=chunk)
                response.raise_for_status()

                uploaded = chunk_end
                if progress_callback:
                    progress_callback(uploaded, total_size)

            return response.json()

    async def upload_batch(
        self,
        site_id: str,
        drive_id: str,
        source_dir: str | Path,
        destination_folder: str = "",
        max_concurrent: int = 5,
        progress_callback: Callable[[str, int, int], None] | None = None,
    ) -> list[dict]:
        """Faz upload de múltiplos arquivos em paralelo."""
        source_dir = Path(source_dir)
        files = [f for f in source_dir.rglob("*") if f.is_file()]

        semaphore = asyncio.Semaphore(max_concurrent)
        uploaded = []
        completed = 0

        async def upload_one(local_file: Path) -> dict:
            nonlocal completed
            async with semaphore:
                relative_path = local_file.relative_to(source_dir)
                remote_path = (
                    f"{destination_folder}/{relative_path}"
                    if destination_folder
                    else str(relative_path)
                )

                result = await self.upload(site_id, drive_id, remote_path, local_file)

                completed += 1
                if progress_callback:
                    progress_callback(local_file.name, completed, len(files))

                return result

        tasks = [upload_one(f) for f in files]
        uploaded = await asyncio.gather(*tasks)

        return list(uploaded)

    # =========================================================================
    # Criar Pasta
    # =========================================================================

    async def create_folder(
        self,
        site_id: str,
        drive_id: str,
        folder_path: str,
    ) -> dict:
        """Cria uma pasta."""
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
            response = await self._request("POST", url, json={
                "name": folder_name,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "fail"
            })

            if response.status_code == 409:
                return await self.get_file_metadata(site_id, drive_id, folder_path)

            response.raise_for_status()
            return response.json()
        except httpx.HTTPError as e:
            raise FolderCreateError(f"Erro ao criar pasta: {e}")

    # =========================================================================
    # Deletar
    # =========================================================================

    async def delete(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
    ) -> bool:
        """Deleta um arquivo ou pasta."""
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/root:/{file_path}"

        try:
            response = await self._request("DELETE", url)

            if response.status_code == 404:
                raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

            if response.status_code == 204:
                return True

            response.raise_for_status()
            return True
        except httpx.HTTPError as e:
            raise DeleteError(f"Erro ao deletar: {e}")

    # =========================================================================
    # Mover e Copiar
    # =========================================================================

    async def move(
        self,
        site_id: str,
        drive_id: str,
        source_path: str,
        destination_folder: str,
        new_name: str | None = None,
    ) -> dict:
        """Move um arquivo ou pasta."""
        item = await self.get_file_metadata(site_id, drive_id, source_path)
        item_id = item["id"]

        if destination_folder:
            dest = await self.get_file_metadata(site_id, drive_id, destination_folder)
            parent_ref = {"id": dest["id"]}
        else:
            parent_ref = {"path": f"/drives/{drive_id}/root"}

        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/items/{item_id}"

        body = {"parentReference": parent_ref}
        if new_name:
            body["name"] = new_name

        try:
            response = await self._request("PATCH", url, json=body)
            response.raise_for_status()
            return response.json()
        except httpx.HTTPError as e:
            raise MoveError(f"Erro ao mover: {e}")

    async def copy(
        self,
        site_id: str,
        drive_id: str,
        source_path: str,
        destination_folder: str,
        new_name: str | None = None,
    ) -> str:
        """Copia um arquivo ou pasta."""
        item = await self.get_file_metadata(site_id, drive_id, source_path)
        item_id = item["id"]

        dest = await self.get_file_metadata(site_id, drive_id, destination_folder)

        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/items/{item_id}/copy"

        body = {
            "parentReference": {"driveId": drive_id, "id": dest["id"]},
            "name": new_name or item["name"],
        }

        try:
            response = await self._request("POST", url, json=body)

            if response.status_code == 202:
                return response.headers.get("Location", "")

            response.raise_for_status()
            return ""
        except httpx.HTTPError as e:
            raise MoveError(f"Erro ao copiar: {e}")

    # =========================================================================
    # Compartilhamento
    # =========================================================================

    async def create_share_link(
        self,
        site_id: str,
        drive_id: str,
        file_path: str,
        link_type: str = "view",
        scope: str = "anonymous",
        expiration: str | None = None,
    ) -> dict:
        """Cria um link de compartilhamento."""
        item = await self.get_file_metadata(site_id, drive_id, file_path)
        item_id = item["id"]

        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}/items/{item_id}/createLink"

        body = {"type": link_type, "scope": scope}
        if expiration:
            body["expirationDateTime"] = expiration

        try:
            response = await self._request("POST", url, json=body)
            response.raise_for_status()
            return response.json()
        except httpx.HTTPError as e:
            raise ShareError(f"Erro ao criar link: {e}")

    # =========================================================================
    # Listas (SharePoint Lists)
    # =========================================================================

    async def list_lists(self, site_id: str) -> list[dict]:
        """Lista todas as listas de um site."""
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists"
        response = await self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    async def get_list(self, site_id: str, list_name: str) -> dict:
        """Obtém uma lista pelo nome ou ID."""
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_name}"
        response = await self._request("GET", url)

        if response.status_code == 200:
            return response.json()

        lists = await self.list_lists(site_id)
        for lst in lists:
            if lst.get("displayName", "").lower() == list_name.lower():
                return lst
            if lst.get("name", "").lower() == list_name.lower():
                return lst

        raise ListError(f"Lista não encontrada: {list_name}")

    async def get_list_columns(self, site_id: str, list_id: str) -> list[dict]:
        """Obtém as colunas de uma lista."""
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}/columns"
        response = await self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    async def list_items(
        self,
        site_id: str,
        list_id: str,
        expand_fields: bool = True,
        filter_query: str | None = None,
        top: int | None = None,
    ) -> list[dict]:
        """Lista itens de uma lista."""
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}/items"

        params = []
        if expand_fields:
            params.append("$expand=fields")
        if filter_query:
            params.append(f"$filter={filter_query}")
        if top:
            params.append(f"$top={top}")

        if params:
            url += "?" + "&".join(params)

        response = await self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    async def list_all_items(
        self,
        site_id: str,
        list_id: str,
        expand_fields: bool = True,
        filter_query: str | None = None,
    ) -> list[dict]:
        """Lista TODOS os itens com paginação automática."""
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
            response = await self._request("GET", url)
            response.raise_for_status()
            data = response.json()

            all_items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")

        return all_items

    async def get_item(
        self,
        site_id: str,
        list_id: str,
        item_id: str,
        expand_fields: bool = True,
    ) -> dict:
        """Obtém um item específico."""
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}/items/{item_id}"

        if expand_fields:
            url += "?$expand=fields"

        response = await self._request("GET", url)

        if response.status_code == 404:
            raise ListError(f"Item não encontrado: {item_id}")

        response.raise_for_status()
        return response.json()

    async def create_item(
        self,
        site_id: str,
        list_id: str,
        fields: dict,
    ) -> dict:
        """Cria um novo item."""
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}/items"

        try:
            response = await self._request("POST", url, json={"fields": fields})
            response.raise_for_status()
            return response.json()
        except httpx.HTTPError as e:
            raise ListError(f"Erro ao criar item: {e}")

    async def update_item(
        self,
        site_id: str,
        list_id: str,
        item_id: str,
        fields: dict,
    ) -> dict:
        """Atualiza um item."""
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}/items/{item_id}/fields"

        try:
            response = await self._request("PATCH", url, json=fields)
            response.raise_for_status()
            return response.json()
        except httpx.HTTPError as e:
            raise ListError(f"Erro ao atualizar item: {e}")

    async def delete_item(
        self,
        site_id: str,
        list_id: str,
        item_id: str,
    ) -> bool:
        """Deleta um item."""
        url = f"{self.GRAPH_BASE_URL}/sites/{site_id}/lists/{list_id}/items/{item_id}"

        try:
            response = await self._request("DELETE", url)

            if response.status_code == 404:
                raise ListError(f"Item não encontrado: {item_id}")

            return response.status_code == 204

        except httpx.HTTPError as e:
            raise ListError(f"Erro ao deletar item: {e}")

    async def batch_create_items(
        self,
        site_id: str,
        list_id: str,
        items: list[dict],
        max_concurrent: int = 5,
    ) -> list[dict]:
        """Cria múltiplos itens em paralelo."""
        semaphore = asyncio.Semaphore(max_concurrent)

        async def create_one(fields: dict) -> dict:
            async with semaphore:
                return await self.create_item(site_id, list_id, fields)

        tasks = [create_one(f) for f in items]
        return list(await asyncio.gather(*tasks))

    async def batch_update_items(
        self,
        site_id: str,
        list_id: str,
        updates: list[tuple[str, dict]],
        max_concurrent: int = 5,
    ) -> list[dict]:
        """Atualiza múltiplos itens em paralelo."""
        semaphore = asyncio.Semaphore(max_concurrent)

        async def update_one(item_id: str, fields: dict) -> dict:
            async with semaphore:
                return await self.update_item(site_id, list_id, item_id, fields)

        tasks = [update_one(item_id, fields) for item_id, fields in updates]
        return list(await asyncio.gather(*tasks))

    async def batch_delete_items(
        self,
        site_id: str,
        list_id: str,
        item_ids: list[str],
        max_concurrent: int = 5,
    ) -> int:
        """Deleta múltiplos itens em paralelo."""
        semaphore = asyncio.Semaphore(max_concurrent)

        async def delete_one(item_id: str) -> bool:
            async with semaphore:
                return await self.delete_item(site_id, list_id, item_id)

        tasks = [delete_one(item_id) for item_id in item_ids]
        results = await asyncio.gather(*tasks)
        return sum(1 for r in results if r)

    # =========================================================================
    # Search (Busca Global)
    # =========================================================================

    async def search(
        self,
        query: str,
        entity_types: list[str] | None = None,
        site_id: str | None = None,
        size: int = 25,
    ) -> list[dict]:
        """Busca global no SharePoint."""
        url = f"{self.GRAPH_BASE_URL}/search/query"

        requests_body = {
            "queryString": query,
            "size": size,
            "entityTypes": entity_types or ["driveItem", "listItem", "site"],
        }

        if site_id:
            requests_body["query"] = {"queryString": f"{query} AND siteId:{site_id}"}

        response = await self._request("POST", url, json={"requests": [requests_body]})
        response.raise_for_status()

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

    async def search_files(
        self,
        query: str,
        file_extension: str | None = None,
        site_id: str | None = None,
        size: int = 25,
    ) -> list[dict]:
        """Busca arquivos no SharePoint."""
        if file_extension:
            search_query = f"{query} filetype:{file_extension}"
        else:
            search_query = query
        return await self.search(
            search_query, entity_types=["driveItem"], site_id=site_id, size=size
        )

    # =========================================================================
    # Acesso por ID
    # =========================================================================

    async def get_item_by_id(self, drive_id: str, item_id: str) -> dict:
        """Obtém um arquivo ou pasta pelo ID."""
        url = f"{self.GRAPH_BASE_URL}/drives/{drive_id}/items/{item_id}"
        response = await self._request("GET", url)

        if response.status_code == 404:
            raise FileNotFoundError(f"Item não encontrado: {item_id}")

        response.raise_for_status()
        return response.json()

    async def download_by_id(
        self,
        drive_id: str,
        item_id: str,
        destination: str | Path,
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> Path:
        """Baixa um arquivo pelo ID."""
        destination = Path(destination)
        url = f"{self.GRAPH_BASE_URL}/drives/{drive_id}/items/{item_id}/content"

        async with httpx.AsyncClient(follow_redirects=True, timeout=120.0) as client:
            async with client.stream("GET", url, headers=self._get_headers()) as response:
                if response.status_code == 404:
                    raise FileNotFoundError(f"Item não encontrado: {item_id}")

                response.raise_for_status()

                destination.parent.mkdir(parents=True, exist_ok=True)
                total_size = int(response.headers.get("content-length", 0))
                downloaded = 0

                with open(destination, "wb") as f:
                    async for chunk in response.aiter_bytes(chunk_size=8192):
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_callback and total_size:
                            progress_callback(downloaded, total_size)

        return destination

    async def delete_by_id(self, drive_id: str, item_id: str) -> bool:
        """Deleta um arquivo ou pasta pelo ID."""
        url = f"{self.GRAPH_BASE_URL}/drives/{drive_id}/items/{item_id}"
        response = await self._request("DELETE", url)

        if response.status_code == 404:
            raise FileNotFoundError(f"Item não encontrado: {item_id}")

        return response.status_code == 204

    # =========================================================================
    # Microsoft Teams
    # =========================================================================

    async def list_teams(self) -> list[dict]:
        """Lista todos os times do Microsoft Teams."""
        url = f"{self.GRAPH_BASE_URL}/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"
        response = await self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    async def get_team(self, team_id: str) -> dict:
        """Obtém informações de um time."""
        url = f"{self.GRAPH_BASE_URL}/teams/{team_id}"
        response = await self._request("GET", url)
        response.raise_for_status()
        return response.json()

    async def list_team_channels(self, team_id: str) -> list[dict]:
        """Lista canais de um time."""
        url = f"{self.GRAPH_BASE_URL}/teams/{team_id}/channels"
        response = await self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    async def get_team_drive(self, team_id: str) -> dict:
        """Obtém o drive de um time."""
        url = f"{self.GRAPH_BASE_URL}/groups/{team_id}/drive"
        response = await self._request("GET", url)
        response.raise_for_status()
        return response.json()

    async def list_team_files(self, team_id: str, folder_path: str = "") -> list[dict]:
        """Lista arquivos de um time."""
        drive = await self.get_team_drive(team_id)
        drive_id = drive["id"]

        if folder_path:
            url = f"{self.GRAPH_BASE_URL}/drives/{drive_id}/root:/{folder_path}:/children"
        else:
            url = f"{self.GRAPH_BASE_URL}/drives/{drive_id}/root/children"

        response = await self._request("GET", url)
        response.raise_for_status()
        return response.json().get("value", [])

    async def download_team_file(
        self,
        team_id: str,
        file_path: str,
        destination: str | Path,
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> Path:
        """Baixa um arquivo de um time."""
        drive = await self.get_team_drive(team_id)
        drive_id = drive["id"]

        destination = Path(destination)
        url = f"{self.GRAPH_BASE_URL}/drives/{drive_id}/root:/{file_path}:/content"

        async with httpx.AsyncClient(follow_redirects=True, timeout=120.0) as client:
            async with client.stream("GET", url, headers=self._get_headers()) as response:
                if response.status_code == 404:
                    raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

                response.raise_for_status()

                destination.parent.mkdir(parents=True, exist_ok=True)
                total_size = int(response.headers.get("content-length", 0))
                downloaded = 0

                with open(destination, "wb") as f:
                    async for chunk in response.aiter_bytes(chunk_size=8192):
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_callback and total_size:
                            progress_callback(downloaded, total_size)

        return destination

    # =========================================================================
    # Métodos de conveniência
    # =========================================================================

    async def download_file(
        self,
        hostname: str,
        site_path: str,
        file_path: str,
        destination: str | Path,
        drive_name: str = "Documents",
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> Path:
        """Baixa um arquivo (método simplificado)."""
        site = await self.get_site(hostname, site_path)
        drive = await self.get_drive(site["id"], drive_name)
        return await self.download(
            site["id"], drive["id"], file_path, destination, progress_callback
        )

    async def upload_file(
        self,
        hostname: str,
        site_path: str,
        file_path: str,
        source: str | Path,
        drive_name: str = "Documents",
        progress_callback: Callable[[int, int], None] | None = None,
    ) -> dict:
        """Faz upload de um arquivo (método simplificado)."""
        site = await self.get_site(hostname, site_path)
        drive = await self.get_drive(site["id"], drive_name)
        return await self.upload(site["id"], drive["id"], file_path, source, progress_callback)
