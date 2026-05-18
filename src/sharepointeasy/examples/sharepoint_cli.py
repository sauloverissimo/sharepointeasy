"""SharePoint CLI — template genérico para acessar pastas oficiais e fazer pequenos ajustes.

================================================================================
COMO USAR ESTE TEMPLATE EM UM NOVO PROJETO
================================================================================

1) Copie este arquivo (sharepoint_cli.py) para a raiz do seu projeto.

2) Crie um arquivo .env na mesma pasta com as credenciais Microsoft Graph:

    MICROSOFT_CLIENTE_ID=...
    MICROSOFT_CLIENTE_SECRET=...
    MICROSOFT_TENANT_ID=...

   (Os nomes MICROSOFT_CLIENT_ID / MICROSOFT_CLIENT_SECRET também são aceitos.)

3) Crie um virtualenv e instale a sharepointeasy (lib do Saulo) + dependências:

    python3 -m venv .venv
    .venv/bin/pip install --upgrade pip
    .venv/bin/pip install git+https://github.com/sauloverissimo/sharepointeasy.git python-dotenv

4) Customize a configuração ABAIXO (bloco CONFIG) para apontar para o seu site,
   drive e pasta base no SharePoint. Verifique a URL do site no navegador:

    https://<HOSTNAME>/<SITE_PATH>/...
    Ex: https://contoso.sharepoint.com/sites/MySite
         → SHAREPOINT_HOSTNAME   = "contoso.sharepoint.com"
         → SHAREPOINT_SITE_PATH  = "sites/MySite"

   Para a pasta base, navegue até ela no SP e pegue o path a partir do drive:

    Documentos > MinhaPasta > MinhaSubpasta
         → SHAREPOINT_BASE_FOLDER = "MinhaPasta/MinhaSubpasta"
         → DRIVE_NAME             = "Documents"   (nome interno do drive default)

5) Comandos disponíveis:

    .venv/bin/python sharepoint_cli.py info               # mostra config atual
    .venv/bin/python sharepoint_cli.py list               # lista arquivos da pasta base
    .venv/bin/python sharepoint_cli.py ls <subpasta>      # lista uma subpasta
    .venv/bin/python sharepoint_cli.py pull <nome>        # baixa arquivo → workspace local
    .venv/bin/python sharepoint_cli.py pull-all           # baixa todos
    .venv/bin/python sharepoint_cli.py push <nome>        # sobe arquivo local (sobrescreve oficial)
    .venv/bin/python sharepoint_cli.py push-all           # sobe todos os arquivos locais

6) Fluxo típico para ajustar um arquivo oficial:

    a) pull <nome>           # baixa versão atual do SP para sp_workspace/
    b) (edite localmente em sp_workspace/<nome> com Word ou via python-docx)
    c) push <nome>           # sobe de volta, sobrescrevendo o oficial

   Se o SP retornar erro 423 (Locked), feche o arquivo no Word/Word Online,
   aguarde ~30s e tente o push novamente.

================================================================================
"""

import os
import sys
from pathlib import Path
from dotenv import load_dotenv
from sharepointeasy import SharePointClient, create_progress_callback


# ============================================================================
# CONFIG — Customize para cada projeto/cliente
# ============================================================================

# URL do site SharePoint (sem "https://")
SHAREPOINT_HOSTNAME = "contoso.sharepoint.com"

# Path do site após o hostname (sem barras no início/fim)
SHAREPOINT_SITE_PATH = "sites/MySite"

# Pasta base dentro do drive (path completo a partir da raiz do drive)
# Use "/" para separar níveis. Acentos/cedilhas são suportados.
SHAREPOINT_BASE_FOLDER = "MinhaPasta/MinhaSubpasta"

# Nome interno do drive. Para a biblioteca padrão "Documentos" use "Documents".
# Para outras bibliotecas, use o nome exato como aparece no SP.
DRIVE_NAME = "Documents"

# Pasta local onde os arquivos baixados ficarão.
# Sugestão: usar "sp_workspace" para não conflitar com outras pastas do projeto.
LOCAL_DIR_NAME = "sp_workspace"

# Caminho do .env (por padrão, ao lado deste arquivo)
ENV_PATH = Path(__file__).parent / ".env"

# ============================================================================


PROJECT_ROOT = Path(__file__).parent
LOCAL_DIR = PROJECT_ROOT / LOCAL_DIR_NAME


def connect():
    """Carrega credenciais do .env e conecta ao SharePoint."""
    load_dotenv(ENV_PATH)
    client_id = os.getenv("MICROSOFT_CLIENTE_ID") or os.getenv("MICROSOFT_CLIENT_ID")
    client_secret = os.getenv("MICROSOFT_CLIENTE_SECRET") or os.getenv("MICROSOFT_CLIENT_SECRET")
    tenant_id = os.getenv("MICROSOFT_TENANT_ID")
    if not all([client_id, client_secret, tenant_id]):
        raise SystemExit(
            "ERRO: credenciais Microsoft ausentes no .env\n"
            "      Verifique MICROSOFT_CLIENTE_ID, MICROSOFT_CLIENTE_SECRET, MICROSOFT_TENANT_ID."
        )
    client = SharePointClient(client_id=client_id, client_secret=client_secret, tenant_id=tenant_id)
    site = client.get_site(SHAREPOINT_HOSTNAME, SHAREPOINT_SITE_PATH)
    drive = client.get_drive(site["id"], DRIVE_NAME)
    return client, site["id"], drive["id"]


def _resolve_folder(subpath=None):
    """Combina BASE_FOLDER + subpath opcional."""
    if subpath:
        return f"{SHAREPOINT_BASE_FOLDER}/{subpath.strip('/')}"
    return SHAREPOINT_BASE_FOLDER


def cmd_info():
    """Mostra configuração atual."""
    print(f"Hostname:    {SHAREPOINT_HOSTNAME}")
    print(f"Site path:   {SHAREPOINT_SITE_PATH}")
    print(f"Drive:       {DRIVE_NAME}")
    print(f"Base folder: {SHAREPOINT_BASE_FOLDER}")
    print(f"Workspace:   {LOCAL_DIR}")
    print(f".env:        {ENV_PATH} ({'OK' if ENV_PATH.exists() else 'AUSENTE'})")


def cmd_list(client, site_id, drive_id, subpath=None):
    """Lista arquivos da pasta base (ou de uma subpasta opcional)."""
    folder = _resolve_folder(subpath)
    print(f"Pasta remota: {folder}")
    try:
        files = client.list_files(site_id, drive_id, folder)
    except Exception as e:
        print(f"Erro: {e}")
        return
    if not files:
        print("(pasta vazia)")
        return
    for f in files:
        kind = "[DIR]" if "folder" in f else "[FILE]"
        size = f.get("size", 0)
        mod = f.get("lastModifiedDateTime", "?")
        print(f"  {kind:6s} {f['name']:60s} {size:>12,} bytes  {mod}")


def cmd_pull(client, site_id, drive_id, name):
    """Baixa um arquivo da pasta base para LOCAL_DIR."""
    LOCAL_DIR.mkdir(exist_ok=True)
    dest = LOCAL_DIR / name
    remote_path = f"{SHAREPOINT_BASE_FOLDER}/{name}"
    print(f"⬇ {name}")
    progress = create_progress_callback(f"Download {name}")
    client.download(site_id, drive_id, remote_path, str(dest), progress_callback=progress)
    print(f"  Salvo em: {dest}")


def cmd_pull_all(client, site_id, drive_id):
    """Baixa todos os arquivos da pasta base."""
    files = client.list_files(site_id, drive_id, SHAREPOINT_BASE_FOLDER)
    count = 0
    for f in files:
        if "folder" not in f:
            cmd_pull(client, site_id, drive_id, f["name"])
            count += 1
    print(f"\nTotal: {count} arquivos baixados.")


def cmd_push(client, site_id, drive_id, name):
    """Sobe um arquivo de LOCAL_DIR para a pasta base (sobrescreve)."""
    src = LOCAL_DIR / name
    if not src.exists():
        raise SystemExit(f"ERRO: arquivo local não encontrado: {src}")
    remote_path = f"{SHAREPOINT_BASE_FOLDER}/{name}"
    print(f"⬆ {name}")
    progress = create_progress_callback(f"Upload {name}")
    client.upload(site_id, drive_id, remote_path, str(src), progress_callback=progress)
    print(f"  Subido para: {remote_path}")


def cmd_push_all(client, site_id, drive_id):
    """Sobe todos os arquivos de LOCAL_DIR para a pasta base."""
    if not LOCAL_DIR.exists():
        raise SystemExit(f"ERRO: pasta local não existe: {LOCAL_DIR}")
    count = 0
    for f in LOCAL_DIR.iterdir():
        if f.is_file() and not f.name.startswith("."):
            cmd_push(client, site_id, drive_id, f.name)
            count += 1
    print(f"\nTotal: {count} arquivos enviados.")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return

    cmd = sys.argv[1]

    # 'info' não exige conexão — útil para validar config antes de tentar conectar
    if cmd == "info":
        cmd_info()
        return

    client, site_id, drive_id = connect()

    if cmd == "list":
        cmd_list(client, site_id, drive_id)
    elif cmd == "ls":
        if len(sys.argv) < 3:
            raise SystemExit("Uso: ls <subpasta>")
        cmd_list(client, site_id, drive_id, sys.argv[2])
    elif cmd == "pull":
        if len(sys.argv) < 3:
            raise SystemExit("Uso: pull <nome>")
        cmd_pull(client, site_id, drive_id, sys.argv[2])
    elif cmd == "pull-all":
        cmd_pull_all(client, site_id, drive_id)
    elif cmd == "push":
        if len(sys.argv) < 3:
            raise SystemExit("Uso: push <nome>")
        cmd_push(client, site_id, drive_id, sys.argv[2])
    elif cmd == "push-all":
        cmd_push_all(client, site_id, drive_id)
    else:
        print(f"Comando desconhecido: {cmd}\n")
        print(__doc__)


if __name__ == "__main__":
    main()
