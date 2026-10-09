"""Small process and thread safe primitives for profile-local state."""
from __future__ import annotations

from contextlib import contextmanager
import fcntl
import json
import os
from pathlib import Path
import tempfile
import threading
import time

_locks: dict[str, threading.RLock] = {}
_guard = threading.Lock()
_local = threading.local()


@contextmanager
def file_lock(path: Path, timeout: float = 150):
    path = Path(path)
    key = str(path.absolute())
    with _guard:
        lock = _locks.setdefault(key, threading.RLock())
    deadline = time.monotonic() + timeout
    if not lock.acquire(timeout=timeout):
        raise TimeoutError(f"Timed out waiting for state lock: {path.name}")
    depths = getattr(_local, "depths", {})
    _local.depths = depths
    handle = None
    try:
        if not depths.get(key):
            path.parent.mkdir(parents=True, exist_ok=True)
            handle = path.open("a+")
            os.chmod(path, 0o600)
            while True:
                try:
                    fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                    break
                except BlockingIOError:
                    if time.monotonic() >= deadline:
                        raise TimeoutError(f"Timed out waiting for state lock: {path.name}")
                    time.sleep(0.05)
        depths[key] = depths.get(key, 0) + 1
        try:
            yield
        finally:
            depths[key] -= 1
    finally:
        if handle:
            fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
            handle.close()
        lock.release()


def atomic_write_json(path: Path, data: object) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=f".{path.name}.", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
