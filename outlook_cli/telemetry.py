"""Opt-in, process-local timings. Never collect URLs, tokens, or message content."""
from __future__ import annotations

import threading
import time
from collections import Counter
from contextlib import contextmanager
from dataclasses import dataclass, field


@dataclass
class Profile:
    started: float = field(default_factory=time.perf_counter)
    requests: int = 0
    retries: int = 0
    response_bytes: int = 0
    network_seconds: float = 0.0
    phases: dict = field(default_factory=dict)
    statuses: Counter = field(default_factory=Counter)
    lock: threading.Lock = field(default_factory=threading.Lock, repr=False)


_active: Profile | None = None


def start_profile() -> Profile:
    global _active
    _active = Profile()
    return _active


def record_request(elapsed: float, response=None, retry: bool = False) -> None:
    profile = _active
    if profile is None:
        return
    with profile.lock:
        profile.requests += 1
        profile.retries += int(retry)
        profile.network_seconds += elapsed
        if response is not None:
            profile.statuses[str(response.status_code)] += 1
            profile.response_bytes += len(response.content)


@contextmanager
def timed_phase(name: str):
    start = time.perf_counter()
    try:
        yield
    finally:
        profile = _active
        if profile:
            with profile.lock:
                profile.phases[name] = profile.phases.get(name, 0.0) + time.perf_counter() - start


def finish_profile(profile: Profile) -> dict:
    global _active
    with profile.lock:
        result = {
            "elapsed_ms": round((time.perf_counter() - profile.started) * 1000, 3),
            "request_count": profile.requests, "retry_count": profile.retries,
            "network_ms_sum": round(profile.network_seconds * 1000, 3),
            "response_bytes": profile.response_bytes, "http_statuses": dict(profile.statuses),
            "phase_ms": {k: round(v * 1000, 3) for k, v in profile.phases.items()},
        }
    if _active is profile:
        _active = None
    return result
