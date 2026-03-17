"""Shared helpers for all command modules."""

from __future__ import annotations

import sys
from collections.abc import MutableMapping
from copy import deepcopy
from typing import Iterator

import click

from .. import account as account_service
from ..auth import get_token as auth_get_token, login as auth_login, verify_token
from ..client import OutlookClient
from ..exceptions import AccountError, AuthRequiredError, OutlookCliError, ResourceNotFoundError, TokenExpiredError, error_code_for_exception
from ..formatter import (
    console,
    print_attachments,
    print_calendars,
    print_categories,
    print_contacts,
    print_email,
    print_email_raw,
    print_error,
    print_event_detail,
    print_events,
    print_folders,
    print_inbox,
    print_meeting_suggestions,
    print_people,
    print_success,
    print_whoami,
)
from ..serialization import error_json, save_json, to_json, to_json_envelope


class ConfigProxy(MutableMapping[str, object]):
    """Resolve config lazily for the selected account profile."""

    def __init__(self):
        self._overrides: dict[str, dict[str, object]] = {}

    def _selected_account(self) -> str:
        return get_account_name(allow_missing=False)

    def _data(self) -> dict:
        account_name = self._selected_account()
        data = deepcopy(account_service.load_account_config(account_name))
        overrides = self._overrides.get(account_name)
        if overrides:
            data.update(overrides)
        return data

    def __getitem__(self, key: str) -> object:
        return self._data()[key]

    def __setitem__(self, key: str, value: object) -> None:
        account_name = self._selected_account()
        self._overrides.setdefault(account_name, {})[key] = value

    def __delitem__(self, key: str) -> None:
        account_name = self._selected_account()
        overrides = self._overrides.setdefault(account_name, {})
        del overrides[key]

    def __iter__(self) -> Iterator[str]:
        return iter(self._data())

    def __len__(self) -> int:
        return len(self._data())

    def get(self, key: str, default=None):
        return self._data().get(key, default)


cfg = ConfigProxy()

# Cache client instances per account profile so auto-relogin only invalidates one profile.
_client_cache: dict[str, OutlookClient] = {}


def account_option(fn):
    return click.option(
        "--account",
        "account_name",
        default=None,
        help="Use a specific account profile",
    )(fn)


def _ctx_account_name() -> str | None:
    ctx = click.get_current_context(silent=True)
    if not ctx:
        return None
    for key in ("account_name", "account"):
        value = ctx.params.get(key)
        if value:
            return value
    return None


def get_account_name(account_name: str | None = None, *, allow_missing: bool = False) -> str:
    selected = account_name or _ctx_account_name()
    return account_service.resolve_account_name(selected, allow_missing=allow_missing)


def get_token(account_name: str | None = None) -> str:
    return auth_get_token(get_account_name(account_name))


def do_login(
    force: bool = False,
    debug: bool = False,
    account_name: str | None = None,
    allow_create: bool = False,
) -> str:
    selected = get_account_name(account_name, allow_missing=allow_create)
    return auth_login(
        force=force,
        debug=debug,
        account_name=selected,
        allow_create=allow_create,
    )


def _get_client(account_name: str | None = None) -> OutlookClient:
    selected = get_account_name(account_name)
    if selected not in _client_cache:
        try:
            token = get_token()
        except (AuthRequiredError, RuntimeError, AccountError) as exc:
            print_error(str(exc))
            sys.exit(1)
        _client_cache[selected] = OutlookClient(token, account_name=selected)
        account_service.touch_account(selected)
    return _client_cache[selected]


def _is_piped() -> bool:
    """True when stdout is not a terminal (piped to another command or file)."""
    return not sys.stdout.isatty()


def _wants_json(as_json: bool) -> bool:
    """True if JSON output is needed: explicit --json flag OR piped stdout."""
    return as_json or _is_piped()


def _is_json_mode() -> bool:
    """Check JSON mode from Click context (used by error handler)."""
    ctx = click.get_current_context(silent=True)
    explicit = bool(ctx and ctx.params.get("as_json"))
    return explicit or _is_piped()


def _handle_api_error(fn):
    """Decorator to catch common API errors. Auto re-login on 401."""
    import functools

    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except TokenExpiredError:
            selected = get_account_name()
            if _is_json_mode():
                click.echo(error_json("session_expired", "Token expired. Attempting re-login..."))
            else:
                print_error("Token expired. Attempting re-login...")
            try:
                login_kwargs = {"account_name": selected} if selected != "default" or _ctx_account_name() else {}
                do_login(**login_kwargs)
                print_success("Re-login successful. Retrying...")
                _client_cache.pop(selected, None)
                return fn(*args, **kwargs)
            except Exception:
                if _is_json_mode():
                    click.echo(error_json("auth_failed", "Auto re-login failed. Run: outlook login --force"))
                else:
                    print_error("Auto re-login failed. Run: outlook login --force")
                sys.exit(1)
        except OutlookCliError as exc:
            if _is_json_mode():
                click.echo(error_json(error_code_for_exception(exc), str(exc)))
            else:
                print_error(str(exc))
            sys.exit(1)
        except Exception as exc:
            if _is_json_mode():
                click.echo(error_json("unknown_error", str(exc)))
            else:
                print_error(f"Error: {exc}")
            sys.exit(1)

    return wrapper
