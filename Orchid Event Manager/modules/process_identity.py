from __future__ import annotations

import os
from typing import Callable, Any


def process_is_running(pid: int) -> bool:
    """Return whether a process ID still exists without changing it."""
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
        return True
    except (ProcessLookupError, PermissionError, OSError):
        return False


def product_master_process_matches(pid: int, command_reader: Callable[[int], str]) -> bool:
    """Return whether an existing PID is Orchid's Product Master process.

    The command-line marker is deliberate: a stale PID file can otherwise
    point to an unrelated process after macOS recycles its numeric PID.
    ``command_reader`` keeps this small safety check independently testable.
    """
    if not process_is_running(pid):
        return False
    try:
        return "--product-master" in str(command_reader(pid) or "")
    except Exception:
        return False
