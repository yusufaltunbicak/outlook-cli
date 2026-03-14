"""Shared helpers for all command modules."""

from __future__ import annotations

import sys

import click

from ..auth import get_token, login as do_login, verify_token
from ..client import OutlookClient
from ..config import load_config
from ..exceptions import AuthRequiredError, OutlookCliError, ResourceNotFoundError, TokenExpiredError, error_code_for_exception
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

cfg = load_config()

# Cache client instance per session so auto-relogin can invalidate it
_client_cache: dict[str, OutlookClient] = {}


def _get_client() -> OutlookClient:
    if "c" not in _client_cache:
        try:
            token = get_token()
        except (AuthRequiredError, RuntimeError) as e:
            print_error(str(e))
            sys.exit(1)
        _client_cache["c"] = OutlookClient(token)
    return _client_cache["c"]


def _is_json_mode() -> bool:
    """Check if current command was invoked with --json flag."""
    ctx = click.get_current_context(silent=True)
    return bool(ctx and ctx.params.get("as_json"))


def _handle_api_error(fn):
    """Decorator to catch common API errors. Auto re-login on 401."""
    import functools

    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except TokenExpiredError:
            if _is_json_mode():
                click.echo(error_json("session_expired", "Token expired. Attempting re-login..."))
            else:
                print_error("Token expired. Attempting re-login...")
            try:
                token = do_login()
                print_success("Re-login successful. Retrying...")
                _client_cache.clear()
                return fn(*args, **kwargs)
            except Exception:
                if _is_json_mode():
                    click.echo(error_json("auth_failed", "Auto re-login failed. Run: outlook login --force"))
                else:
                    print_error("Auto re-login failed. Run: outlook login --force")
                sys.exit(1)
        except OutlookCliError as e:
            if _is_json_mode():
                click.echo(error_json(error_code_for_exception(e), str(e)))
            else:
                print_error(str(e))
            sys.exit(1)
        except Exception as e:
            if _is_json_mode():
                click.echo(error_json("unknown_error", str(e)))
            else:
                print_error(f"Error: {e}")
            sys.exit(1)

    return wrapper
