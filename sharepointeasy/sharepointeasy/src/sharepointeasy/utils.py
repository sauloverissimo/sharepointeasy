"""Utilitários para sharepointeasy."""

import sys
from typing import Callable


def create_progress_callback(
    description: str = "Progress",
    show_percentage: bool = True,
    show_bytes: bool = True,
) -> Callable[[int, int], None]:
    """
    Cria um callback de progresso simples para terminal.

    Args:
        description: Descrição da operação
        show_percentage: Mostrar porcentagem
        show_bytes: Mostrar bytes transferidos

    Returns:
        Função callback para progresso
    """

    def callback(current: int, total: int) -> None:
        if total == 0:
            return

        percentage = (current / total) * 100
        bar_length = 30
        filled = int(bar_length * current / total)
        bar = "█" * filled + "░" * (bar_length - filled)

        parts = [f"\r{description}: [{bar}]"]

        if show_percentage:
            parts.append(f" {percentage:5.1f}%")

        if show_bytes:
            current_mb = current / (1024 * 1024)
            total_mb = total / (1024 * 1024)
            parts.append(f" ({current_mb:.1f}/{total_mb:.1f} MB)")

        sys.stdout.write("".join(parts))
        sys.stdout.flush()

        if current >= total:
            sys.stdout.write("\n")
            sys.stdout.flush()

    return callback


def create_batch_progress_callback(
    description: str = "Files",
) -> Callable[[str, int, int], None]:
    """
    Cria um callback de progresso para operações em lote.

    Args:
        description: Descrição da operação

    Returns:
        Função callback para progresso de batch
    """

    def callback(filename: str, current: int, total: int) -> None:
        bar_length = 20
        filled = int(bar_length * current / total)
        bar = "█" * filled + "░" * (bar_length - filled)

        # Truncar nome do arquivo se muito longo
        max_name_len = 30
        display_name = filename[:max_name_len] + "..." if len(filename) > max_name_len else filename

        sys.stdout.write(f"\r{description}: [{bar}] {current}/{total} - {display_name:<35}")
        sys.stdout.flush()

        if current >= total:
            sys.stdout.write("\n")
            sys.stdout.flush()

    return callback


def format_size(size_bytes: int) -> str:
    """
    Formata tamanho em bytes para formato legível.

    Args:
        size_bytes: Tamanho em bytes

    Returns:
        String formatada (ex: "1.5 MB")
    """
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if abs(size_bytes) < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} PB"


def format_path(path: str, max_length: int = 50) -> str:
    """
    Formata um path para exibição, truncando se necessário.

    Args:
        path: Caminho a formatar
        max_length: Tamanho máximo

    Returns:
        Path formatado
    """
    if len(path) <= max_length:
        return path

    # Manter início e fim
    half = (max_length - 3) // 2
    return f"{path[:half]}...{path[-half:]}"


try:
    from rich.progress import (
        BarColumn,
        DownloadColumn,
        Progress,
        TaskID,
        TextColumn,
        TimeRemainingColumn,
        TransferSpeedColumn,
    )

    def create_rich_progress() -> Progress:
        """
        Cria uma barra de progresso rica usando a biblioteca rich.

        Returns:
            Instância de Progress do rich
        """
        return Progress(
            TextColumn("[bold blue]{task.description}"),
            BarColumn(),
            "[progress.percentage]{task.percentage:>3.1f}%",
            "•",
            DownloadColumn(),
            "•",
            TransferSpeedColumn(),
            "•",
            TimeRemainingColumn(),
        )

    class RichProgressCallback:
        """Wrapper para usar rich Progress com callbacks."""

        def __init__(self, progress: Progress, task_id: TaskID, total: int):
            self.progress = progress
            self.task_id = task_id
            self.total = total
            self._last_current = 0

        def __call__(self, current: int, total: int) -> None:
            advance = current - self._last_current
            self.progress.update(self.task_id, advance=advance)
            self._last_current = current

    RICH_AVAILABLE = True

except ImportError:
    RICH_AVAILABLE = False

    def create_rich_progress():
        raise ImportError("rich is not installed. Install with: pip install rich")

    class RichProgressCallback:
        def __init__(self, *args, **kwargs):
            raise ImportError("rich is not installed. Install with: pip install rich")
