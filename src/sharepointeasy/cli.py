"""Interface de linha de comando para sharepointeasy."""

import argparse
import sys
from pathlib import Path

from .client import SharePointClient
from .exceptions import SharePointError
from .utils import create_batch_progress_callback, create_progress_callback, format_size


def get_client() -> SharePointClient:
    """Cria cliente SharePoint com credenciais do ambiente."""
    try:
        return SharePointClient()
    except SharePointError as e:
        print(f"Erro de autenticação: {e}", file=sys.stderr)
        print("\nConfigure as variáveis de ambiente:", file=sys.stderr)
        print("  export MICROSOFT_CLIENT_ID='...'", file=sys.stderr)
        print("  export MICROSOFT_CLIENT_SECRET='...'", file=sys.stderr)
        print("  export MICROSOFT_TENANT_ID='...'", file=sys.stderr)
        sys.exit(1)


def cmd_list_sites(args: argparse.Namespace) -> None:
    """Lista sites disponíveis."""
    client = get_client()
    sites = client.list_sites()

    if not sites:
        print("Nenhum site encontrado.")
        return

    print(f"\n{'Nome':<30} {'URL':<50}")
    print("-" * 80)
    for site in sites:
        name = site.get("displayName", "N/A")[:30]
        url = site.get("webUrl", "N/A")[:50]
        print(f"{name:<30} {url:<50}")
    print(f"\nTotal: {len(sites)} sites")


def cmd_list_drives(args: argparse.Namespace) -> None:
    """Lista drives de um site."""
    client = get_client()

    # Obter site
    site = client.get_site(args.hostname, args.site_path)
    drives = client.list_drives(site["id"])

    if not drives:
        print("Nenhum drive encontrado.")
        return

    print(f"\n{'Nome':<30} {'ID':<50}")
    print("-" * 80)
    for drive in drives:
        name = drive.get("name", "N/A")[:30]
        drive_id = drive.get("id", "N/A")[:50]
        print(f"{name:<30} {drive_id:<50}")
    print(f"\nTotal: {len(drives)} drives")


def cmd_list_files(args: argparse.Namespace) -> None:
    """Lista arquivos em uma pasta."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    drive = client.get_drive(site["id"], args.drive or "Documents")

    if args.recursive:
        files = client.list_files_recursive(site["id"], drive["id"], args.path or "")
    else:
        files = client.list_files(site["id"], drive["id"], args.path or "")

    if not files:
        print("Nenhum arquivo encontrado.")
        return

    print(f"\n{'Nome':<40} {'Tamanho':<12} {'Modificado':<20}")
    print("-" * 72)
    for item in files:
        name = item.get("name", "N/A")
        if len(name) > 40:
            name = name[:37] + "..."

        size = format_size(item.get("size", 0)) if "size" in item else "pasta"
        modified = item.get("lastModifiedDateTime", "N/A")[:19]

        is_folder = "folder" in item
        prefix = "[DIR]  " if is_folder else "[FILE] "
        print(f"{prefix}{name:<36} {size:<12} {modified:<20}")

    print(f"\nTotal: {len(files)} itens")


def cmd_download(args: argparse.Namespace) -> None:
    """Baixa um arquivo."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    drive = client.get_drive(site["id"], args.drive or "Documents")

    destination = Path(args.destination or Path(args.file_path).name)

    progress = create_progress_callback("Baixando")

    print(f"Baixando: {args.file_path}")
    result = client.download(
        site["id"],
        drive["id"],
        args.file_path,
        destination,
        progress_callback=progress,
    )
    print(f"Salvo em: {result}")


def cmd_download_folder(args: argparse.Namespace) -> None:
    """Baixa uma pasta inteira."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    drive = client.get_drive(site["id"], args.drive or "Documents")

    destination = Path(args.destination or args.folder_path.split("/")[-1])

    progress = create_batch_progress_callback("Baixando")

    print(f"Baixando pasta: {args.folder_path}")
    results = client.download_batch(
        site["id"],
        drive["id"],
        args.folder_path,
        destination,
        progress_callback=progress,
    )
    print(f"\nBaixados {len(results)} arquivos para: {destination}")


def cmd_upload(args: argparse.Namespace) -> None:
    """Faz upload de um arquivo."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    drive = client.get_drive(site["id"], args.drive or "Documents")

    source = Path(args.source)
    if not source.exists():
        print(f"Arquivo não encontrado: {source}", file=sys.stderr)
        sys.exit(1)

    file_path = args.destination or source.name

    progress = create_progress_callback("Enviando")

    print(f"Enviando: {source}")
    result = client.upload(
        site["id"],
        drive["id"],
        file_path,
        source,
        progress_callback=progress,
    )
    print(f"Enviado para: {result.get('webUrl', file_path)}")


def cmd_upload_folder(args: argparse.Namespace) -> None:
    """Faz upload de uma pasta inteira."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    drive = client.get_drive(site["id"], args.drive or "Documents")

    source = Path(args.source)
    if not source.is_dir():
        print(f"Diretório não encontrado: {source}", file=sys.stderr)
        sys.exit(1)

    progress = create_batch_progress_callback("Enviando")

    print(f"Enviando pasta: {source}")
    results = client.upload_batch(
        site["id"],
        drive["id"],
        source,
        args.destination or "",
        progress_callback=progress,
    )
    print(f"\nEnviados {len(results)} arquivos")


def cmd_delete(args: argparse.Namespace) -> None:
    """Deleta um arquivo ou pasta."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    drive = client.get_drive(site["id"], args.drive or "Documents")

    if not args.yes:
        response = input(f"Deletar '{args.path}'? [y/N]: ")
        if response.lower() != "y":
            print("Cancelado.")
            return

    client.delete(site["id"], drive["id"], args.path)
    print(f"Deletado: {args.path}")


def cmd_mkdir(args: argparse.Namespace) -> None:
    """Cria uma pasta."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    drive = client.get_drive(site["id"], args.drive or "Documents")

    if args.parents:
        result = client.create_folder_recursive(site["id"], drive["id"], args.path)
    else:
        result = client.create_folder(site["id"], drive["id"], args.path)

    print(f"Pasta criada: {result.get('webUrl', args.path)}")


def cmd_move(args: argparse.Namespace) -> None:
    """Move um arquivo ou pasta."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    drive = client.get_drive(site["id"], args.drive or "Documents")

    result = client.move(
        site["id"],
        drive["id"],
        args.source,
        args.destination,
        new_name=args.name,
    )
    print(f"Movido para: {result.get('webUrl', args.destination)}")


def cmd_copy(args: argparse.Namespace) -> None:
    """Copia um arquivo ou pasta."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    drive = client.get_drive(site["id"], args.drive or "Documents")

    result = client.copy(
        site["id"],
        drive["id"],
        args.source,
        args.destination,
        new_name=args.name,
    )
    if result:
        print(f"Cópia iniciada. Monitor: {result}")
    else:
        print("Copiado com sucesso.")


def cmd_share(args: argparse.Namespace) -> None:
    """Cria link de compartilhamento."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    drive = client.get_drive(site["id"], args.drive or "Documents")

    result = client.create_share_link(
        site["id"],
        drive["id"],
        args.path,
        link_type=args.type,
        scope=args.scope,
        expiration=args.expiration,
    )

    link = result.get("link", {}).get("webUrl", "N/A")
    print(f"Link de compartilhamento: {link}")


def cmd_versions(args: argparse.Namespace) -> None:
    """Lista versões de um arquivo."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    drive = client.get_drive(site["id"], args.drive or "Documents")

    versions = client.list_versions(site["id"], drive["id"], args.path)

    if not versions:
        print("Nenhuma versão encontrada.")
        return

    print(f"\n{'ID':<20} {'Modificado':<25} {'Tamanho':<12}")
    print("-" * 57)
    for v in versions:
        vid = v.get("id", "N/A")[:20]
        modified = v.get("lastModifiedDateTime", "N/A")[:25]
        size = format_size(v.get("size", 0))
        print(f"{vid:<20} {modified:<25} {size:<12}")


# =========================================================================
# Comandos de Listas
# =========================================================================


def cmd_list_lists(args: argparse.Namespace) -> None:
    """Lista listas do site."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    lists = client.list_lists(site["id"])

    if not lists:
        print("Nenhuma lista encontrada.")
        return

    print(f"\n{'Nome':<35} {'Itens':<10} {'Template':<20}")
    print("-" * 65)
    for lst in lists:
        name = lst.get("displayName", "N/A")[:35]
        item_count = lst.get("list", {}).get("contentTypesEnabled", "?")
        template = lst.get("list", {}).get("template", "N/A")[:20]
        print(f"{name:<35} {str(item_count):<10} {template:<20}")

    print(f"\nTotal: {len(lists)} listas")


def cmd_list_items(args: argparse.Namespace) -> None:
    """Lista itens de uma lista."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    lst = client.get_list(site["id"], args.list_name)

    if args.all:
        items = client.list_all_items(
            site["id"],
            lst["id"],
            filter_query=args.filter,
        )
    else:
        items = client.list_items(
            site["id"],
            lst["id"],
            filter_query=args.filter,
            top=args.top,
        )

    if not items:
        print("Nenhum item encontrado.")
        return

    # Exibir itens
    print(f"\n{'ID':<10} {'Campos':<60}")
    print("-" * 70)
    for item in items:
        item_id = item.get("id", "N/A")[:10]
        fields = item.get("fields", {})

        # Mostrar alguns campos principais
        field_str = ", ".join(
            f"{k}: {str(v)[:15]}"
            for k, v in list(fields.items())[:3]
            if not k.startswith("@")
        )
        print(f"{item_id:<10} {field_str[:60]:<60}")

    print(f"\nTotal: {len(items)} itens")


def cmd_get_item(args: argparse.Namespace) -> None:
    """Obtém detalhes de um item."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    lst = client.get_list(site["id"], args.list_name)
    item = client.get_item(site["id"], lst["id"], args.item_id)

    print(f"\nItem ID: {item.get('id')}")
    print("-" * 40)

    fields = item.get("fields", {})
    for key, value in fields.items():
        if not key.startswith("@"):
            print(f"  {key}: {value}")


def cmd_create_item(args: argparse.Namespace) -> None:
    """Cria um item em uma lista."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    lst = client.get_list(site["id"], args.list_name)

    # Parse fields do formato "Campo=Valor"
    fields = {}
    for field_str in args.fields:
        if "=" in field_str:
            key, value = field_str.split("=", 1)
            fields[key] = value

    if not fields:
        print("Erro: Especifique campos no formato Campo=Valor", file=sys.stderr)
        sys.exit(1)

    result = client.create_item(site["id"], lst["id"], fields)
    print(f"Item criado com ID: {result.get('id')}")


def cmd_update_item(args: argparse.Namespace) -> None:
    """Atualiza um item."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    lst = client.get_list(site["id"], args.list_name)

    # Parse fields
    fields = {}
    for field_str in args.fields:
        if "=" in field_str:
            key, value = field_str.split("=", 1)
            fields[key] = value

    if not fields:
        print("Erro: Especifique campos no formato Campo=Valor", file=sys.stderr)
        sys.exit(1)

    client.update_item(site["id"], lst["id"], args.item_id, fields)
    print(f"Item {args.item_id} atualizado.")


def cmd_delete_item(args: argparse.Namespace) -> None:
    """Deleta um item."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    lst = client.get_list(site["id"], args.list_name)

    if not args.yes:
        response = input(f"Deletar item '{args.item_id}'? [y/N]: ")
        if response.lower() != "y":
            print("Cancelado.")
            return

    client.delete_item(site["id"], lst["id"], args.item_id)
    print(f"Item {args.item_id} deletado.")


def cmd_list_columns(args: argparse.Namespace) -> None:
    """Lista colunas de uma lista."""
    client = get_client()

    site = client.get_site(args.hostname, args.site_path)
    lst = client.get_list(site["id"], args.list_name)
    columns = client.get_list_columns(site["id"], lst["id"])

    if not columns:
        print("Nenhuma coluna encontrada.")
        return

    print(f"\n{'Nome':<25} {'Tipo':<15} {'Obrigatório':<12}")
    print("-" * 52)
    for col in columns:
        name = col.get("displayName", col.get("name", "N/A"))[:25]

        # Detectar tipo
        col_type = "text"
        for t in ["text", "number", "dateTime", "boolean", "choice", "lookup", "person"]:
            if t in col:
                col_type = t
                break

        required = "Sim" if col.get("required") else "Não"
        print(f"{name:<25} {col_type:<15} {required:<12}")


# =========================================================================
# Comandos de Search
# =========================================================================


def cmd_search(args: argparse.Namespace) -> None:
    """Busca global no SharePoint."""
    client = get_client()

    entity_types = None
    if args.type == "files":
        entity_types = ["driveItem"]
    elif args.type == "lists":
        entity_types = ["listItem"]
    elif args.type == "sites":
        entity_types = ["site"]

    results = client.search(
        args.query,
        entity_types=entity_types,
        size=args.limit,
    )

    if not results:
        print("Nenhum resultado encontrado.")
        return

    print(f"\n{'Rank':<6} {'Tipo':<15} {'Nome':<40}")
    print("-" * 61)
    for result in results:
        rank = str(result.get("rank", "-"))[:6]
        resource = result.get("resource", {})
        name = resource.get("name", resource.get("displayName", "N/A"))[:40]

        # Detectar tipo
        res_type = "unknown"
        if "@odata.type" in resource:
            res_type = resource["@odata.type"].split(".")[-1][:15]

        print(f"{rank:<6} {res_type:<15} {name:<40}")

    print(f"\nTotal: {len(results)} resultados")


# =========================================================================
# Comandos de Teams
# =========================================================================


def cmd_list_teams(args: argparse.Namespace) -> None:
    """Lista times do Microsoft Teams."""
    client = get_client()
    teams = client.list_teams()

    if not teams:
        print("Nenhum time encontrado.")
        return

    print(f"\n{'Nome':<35} {'ID':<40}")
    print("-" * 75)
    for team in teams:
        name = team.get("displayName", "N/A")[:35]
        team_id = team.get("id", "N/A")[:40]
        print(f"{name:<35} {team_id:<40}")

    print(f"\nTotal: {len(teams)} times")


def cmd_list_team_channels(args: argparse.Namespace) -> None:
    """Lista canais de um time."""
    client = get_client()
    channels = client.list_team_channels(args.team_id)

    if not channels:
        print("Nenhum canal encontrado.")
        return

    print(f"\n{'Nome':<35} {'ID':<40}")
    print("-" * 75)
    for channel in channels:
        name = channel.get("displayName", "N/A")[:35]
        channel_id = channel.get("id", "N/A")[:40]
        print(f"{name:<35} {channel_id:<40}")


def cmd_list_team_files(args: argparse.Namespace) -> None:
    """Lista arquivos de um time."""
    client = get_client()
    files = client.list_team_files(args.team_id, args.path or "")

    if not files:
        print("Nenhum arquivo encontrado.")
        return

    print(f"\n{'Nome':<40} {'Tamanho':<12} {'Modificado':<20}")
    print("-" * 72)
    for item in files:
        name = item.get("name", "N/A")[:40]
        size = format_size(item.get("size", 0)) if "size" in item else "pasta"
        modified = item.get("lastModifiedDateTime", "N/A")[:19]
        is_folder = "folder" in item
        prefix = "[DIR]  " if is_folder else "[FILE] "
        print(f"{prefix}{name:<36} {size:<12} {modified:<20}")


def cmd_download_team_file(args: argparse.Namespace) -> None:
    """Baixa um arquivo de um time."""
    client = get_client()

    destination = Path(args.destination or Path(args.file_path).name)
    progress = create_progress_callback("Baixando")

    print(f"Baixando: {args.file_path}")
    result = client.download_team_file(
        args.team_id,
        args.file_path,
        destination,
        progress_callback=progress,
    )
    print(f"Salvo em: {result}")


# =========================================================================
# Comandos de Acesso por ID
# =========================================================================


def cmd_download_by_id(args: argparse.Namespace) -> None:
    """Baixa um arquivo pelo ID."""
    client = get_client()

    destination = Path(args.destination or f"file_{args.item_id}")
    progress = create_progress_callback("Baixando")

    print(f"Baixando item ID: {args.item_id}")
    result = client.download_by_id(
        args.drive_id,
        args.item_id,
        destination,
        progress_callback=progress,
    )
    print(f"Salvo em: {result}")


def cmd_get_item_by_id(args: argparse.Namespace) -> None:
    """Obtém informações de um item pelo ID."""
    client = get_client()
    item = client.get_item_by_id(args.drive_id, args.item_id)

    print(f"\nItem: {item.get('name', 'N/A')}")
    print("-" * 40)
    print(f"  ID: {item.get('id')}")
    print(f"  Tamanho: {format_size(item.get('size', 0))}")
    print(f"  Modificado: {item.get('lastModifiedDateTime', 'N/A')}")
    print(f"  Criado: {item.get('createdDateTime', 'N/A')}")
    print(f"  Web URL: {item.get('webUrl', 'N/A')}")


def create_parser() -> argparse.ArgumentParser:
    """Cria parser de argumentos."""
    parser = argparse.ArgumentParser(
        prog="sharepointeasy",
        description="Easy SharePoint file operations",
    )

    # Argumentos globais
    parser.add_argument(
        "--hostname",
        "-H",
        default=None,
        help="SharePoint hostname (ex: contoso.sharepoint.com)",
    )
    parser.add_argument(
        "--site-path",
        "-S",
        default=None,
        help="Site path (ex: sites/MySite)",
    )
    parser.add_argument(
        "--drive",
        "-D",
        default=None,
        help="Drive name (default: Documents)",
    )

    subparsers = parser.add_subparsers(dest="command", help="Comandos disponíveis")

    # list-sites
    sp = subparsers.add_parser("list-sites", help="Lista sites disponíveis")
    sp.set_defaults(func=cmd_list_sites)

    # list-drives
    sp = subparsers.add_parser("list-drives", help="Lista drives de um site")
    sp.set_defaults(func=cmd_list_drives)

    # list / ls
    sp = subparsers.add_parser("list", aliases=["ls"], help="Lista arquivos")
    sp.add_argument("path", nargs="?", default="", help="Caminho da pasta")
    sp.add_argument("-r", "--recursive", action="store_true", help="Listar recursivamente")
    sp.set_defaults(func=cmd_list_files)

    # download
    sp = subparsers.add_parser("download", aliases=["dl"], help="Baixa um arquivo")
    sp.add_argument("file_path", help="Caminho do arquivo no SharePoint")
    sp.add_argument("-o", "--destination", help="Caminho local de destino")
    sp.set_defaults(func=cmd_download)

    # download-folder
    sp = subparsers.add_parser("download-folder", help="Baixa uma pasta inteira")
    sp.add_argument("folder_path", help="Caminho da pasta no SharePoint")
    sp.add_argument("-o", "--destination", help="Diretório local de destino")
    sp.set_defaults(func=cmd_download_folder)

    # upload
    sp = subparsers.add_parser("upload", aliases=["up"], help="Faz upload de um arquivo")
    sp.add_argument("source", help="Arquivo local")
    sp.add_argument("-d", "--destination", help="Caminho de destino no SharePoint")
    sp.set_defaults(func=cmd_upload)

    # upload-folder
    sp = subparsers.add_parser("upload-folder", help="Faz upload de uma pasta inteira")
    sp.add_argument("source", help="Diretório local")
    sp.add_argument("-d", "--destination", help="Pasta de destino no SharePoint")
    sp.set_defaults(func=cmd_upload_folder)

    # delete / rm
    sp = subparsers.add_parser("delete", aliases=["rm"], help="Deleta arquivo ou pasta")
    sp.add_argument("path", help="Caminho a deletar")
    sp.add_argument("-y", "--yes", action="store_true", help="Não pedir confirmação")
    sp.set_defaults(func=cmd_delete)

    # mkdir
    sp = subparsers.add_parser("mkdir", help="Cria uma pasta")
    sp.add_argument("path", help="Caminho da pasta")
    sp.add_argument("-p", "--parents", action="store_true", help="Criar pastas pai")
    sp.set_defaults(func=cmd_mkdir)

    # move / mv
    sp = subparsers.add_parser("move", aliases=["mv"], help="Move arquivo ou pasta")
    sp.add_argument("source", help="Origem")
    sp.add_argument("destination", help="Pasta de destino")
    sp.add_argument("-n", "--name", help="Novo nome")
    sp.set_defaults(func=cmd_move)

    # copy / cp
    sp = subparsers.add_parser("copy", aliases=["cp"], help="Copia arquivo ou pasta")
    sp.add_argument("source", help="Origem")
    sp.add_argument("destination", help="Pasta de destino")
    sp.add_argument("-n", "--name", help="Novo nome")
    sp.set_defaults(func=cmd_copy)

    # share
    sp = subparsers.add_parser("share", help="Cria link de compartilhamento")
    sp.add_argument("path", help="Caminho do arquivo")
    sp.add_argument(
        "-t", "--type",
        choices=["view", "edit", "embed"],
        default="view",
        help="Tipo de permissão",
    )
    sp.add_argument(
        "-s", "--scope",
        choices=["anonymous", "organization"],
        default="anonymous",
        help="Escopo do link",
    )
    sp.add_argument("-e", "--expiration", help="Data de expiração (ISO 8601)")
    sp.set_defaults(func=cmd_share)

    # versions
    sp = subparsers.add_parser("versions", help="Lista versões de um arquivo")
    sp.add_argument("path", help="Caminho do arquivo")
    sp.set_defaults(func=cmd_versions)

    # =========================================================================
    # Comandos de Listas
    # =========================================================================

    # list-lists
    sp = subparsers.add_parser("list-lists", help="Lista listas do site")
    sp.set_defaults(func=cmd_list_lists)

    # list-items
    sp = subparsers.add_parser("list-items", help="Lista itens de uma lista")
    sp.add_argument("list_name", help="Nome ou ID da lista")
    sp.add_argument("-f", "--filter", help="Filtro OData (ex: fields/Status eq 'Ativo')")
    sp.add_argument("-t", "--top", type=int, help="Número máximo de itens")
    sp.add_argument("-a", "--all", action="store_true", help="Listar todos (paginação automática)")
    sp.set_defaults(func=cmd_list_items)

    # get-item
    sp = subparsers.add_parser("get-item", help="Obtém detalhes de um item")
    sp.add_argument("list_name", help="Nome ou ID da lista")
    sp.add_argument("item_id", help="ID do item")
    sp.set_defaults(func=cmd_get_item)

    # create-item
    sp = subparsers.add_parser("create-item", help="Cria um item em uma lista")
    sp.add_argument("list_name", help="Nome ou ID da lista")
    sp.add_argument("fields", nargs="+", help="Campos no formato Campo=Valor")
    sp.set_defaults(func=cmd_create_item)

    # update-item
    sp = subparsers.add_parser("update-item", help="Atualiza um item")
    sp.add_argument("list_name", help="Nome ou ID da lista")
    sp.add_argument("item_id", help="ID do item")
    sp.add_argument("fields", nargs="+", help="Campos no formato Campo=Valor")
    sp.set_defaults(func=cmd_update_item)

    # delete-item
    sp = subparsers.add_parser("delete-item", help="Deleta um item de uma lista")
    sp.add_argument("list_name", help="Nome ou ID da lista")
    sp.add_argument("item_id", help="ID do item")
    sp.add_argument("-y", "--yes", action="store_true", help="Não pedir confirmação")
    sp.set_defaults(func=cmd_delete_item)

    # list-columns
    sp = subparsers.add_parser("list-columns", help="Lista colunas de uma lista")
    sp.add_argument("list_name", help="Nome ou ID da lista")
    sp.set_defaults(func=cmd_list_columns)

    # =========================================================================
    # Comandos de Search
    # =========================================================================

    # search
    sp = subparsers.add_parser("search", help="Busca global no SharePoint")
    sp.add_argument("query", help="Texto de busca")
    sp.add_argument(
        "-t", "--type",
        choices=["all", "files", "lists", "sites"],
        default="all",
        help="Tipo de entidade",
    )
    sp.add_argument("-l", "--limit", type=int, default=25, help="Número máximo de resultados")
    sp.set_defaults(func=cmd_search)

    # =========================================================================
    # Comandos de Teams
    # =========================================================================

    # list-teams
    sp = subparsers.add_parser("list-teams", help="Lista times do Microsoft Teams")
    sp.set_defaults(func=cmd_list_teams)

    # team-channels
    sp = subparsers.add_parser("team-channels", help="Lista canais de um time")
    sp.add_argument("team_id", help="ID do time")
    sp.set_defaults(func=cmd_list_team_channels)

    # team-files
    sp = subparsers.add_parser("team-files", help="Lista arquivos de um time")
    sp.add_argument("team_id", help="ID do time")
    sp.add_argument("path", nargs="?", default="", help="Caminho da pasta")
    sp.set_defaults(func=cmd_list_team_files)

    # team-download
    sp = subparsers.add_parser("team-download", help="Baixa arquivo de um time")
    sp.add_argument("team_id", help="ID do time")
    sp.add_argument("file_path", help="Caminho do arquivo")
    sp.add_argument("-o", "--destination", help="Caminho local de destino")
    sp.set_defaults(func=cmd_download_team_file)

    # =========================================================================
    # Comandos de Acesso por ID
    # =========================================================================

    # get-by-id
    sp = subparsers.add_parser("get-by-id", help="Obtém informações de um item pelo ID")
    sp.add_argument("drive_id", help="ID do drive")
    sp.add_argument("item_id", help="ID do item")
    sp.set_defaults(func=cmd_get_item_by_id)

    # download-by-id
    sp = subparsers.add_parser("download-by-id", help="Baixa arquivo pelo ID")
    sp.add_argument("drive_id", help="ID do drive")
    sp.add_argument("item_id", help="ID do item")
    sp.add_argument("-o", "--destination", help="Caminho local de destino")
    sp.set_defaults(func=cmd_download_by_id)

    return parser


def main() -> None:
    """Ponto de entrada da CLI."""
    parser = create_parser()
    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(0)

    # Validar hostname e site_path para comandos que precisam
    no_site_commands = [
        "list-sites", "list-teams", "team-channels", "team-files",
        "team-download", "get-by-id", "download-by-id", "search",
    ]
    needs_site = args.command not in no_site_commands
    if needs_site and (not args.hostname or not args.site_path):
        print("Erro: --hostname e --site-path são obrigatórios", file=sys.stderr)
        print("\nExemplo:", file=sys.stderr)
        print("  sharepointeasy -H contoso.sharepoint.com -S sites/MySite list", file=sys.stderr)
        sys.exit(1)

    try:
        args.func(args)
    except SharePointError as e:
        print(f"Erro: {e}", file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        print("\nCancelado pelo usuário.")
        sys.exit(130)


if __name__ == "__main__":
    main()
