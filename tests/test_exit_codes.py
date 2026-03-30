from __future__ import annotations

import json

import click
import httpx

from outlook_cli.commands import _common as common
from outlook_cli.exceptions import (
    EXIT_CODE_AUTH_REQUIRED,
    EXIT_CODE_CONFIG,
    EXIT_CODE_FAILURE,
    EXIT_CODE_NOT_FOUND,
    EXIT_CODE_RATE_LIMITED,
    EXIT_CODE_RETRYABLE,
    EXIT_CODE_USAGE,
    AccountError,
    AuthRequiredError,
    RateLimitError,
    ResourceNotFoundError,
    TokenExpiredError,
    error_code_for_exception,
    exit_code_for_exception,
)


def _http_status_error(status_code: int) -> httpx.HTTPStatusError:
    request = httpx.Request("GET", "https://example.com")
    response = httpx.Response(status_code, request=request)
    return httpx.HTTPStatusError(f"http {status_code}", request=request, response=response)


def test_exit_code_for_exception_known_mappings():
    assert exit_code_for_exception(TokenExpiredError("x")) == EXIT_CODE_AUTH_REQUIRED
    assert exit_code_for_exception(AuthRequiredError("x")) == EXIT_CODE_AUTH_REQUIRED
    assert exit_code_for_exception(ResourceNotFoundError("x")) == EXIT_CODE_NOT_FOUND
    assert exit_code_for_exception(RateLimitError("x")) == EXIT_CODE_RATE_LIMITED
    assert exit_code_for_exception(AccountError("x")) == EXIT_CODE_CONFIG
    assert exit_code_for_exception(click.BadParameter("bad")) == EXIT_CODE_USAGE


def test_exit_code_for_exception_httpx_mappings():
    assert exit_code_for_exception(httpx.ReadTimeout("timeout")) == EXIT_CODE_RETRYABLE
    assert exit_code_for_exception(_http_status_error(404)) == EXIT_CODE_NOT_FOUND
    assert exit_code_for_exception(_http_status_error(429)) == EXIT_CODE_RATE_LIMITED
    assert exit_code_for_exception(_http_status_error(503)) == EXIT_CODE_RETRYABLE
    assert exit_code_for_exception(RuntimeError("boom")) == EXIT_CODE_FAILURE


def test_error_code_for_exception_httpx_mappings():
    assert error_code_for_exception(_http_status_error(404)) == "not_found"
    assert error_code_for_exception(_http_status_error(429)) == "rate_limited"
    assert error_code_for_exception(_http_status_error(503)) == "retryable_error"
    assert error_code_for_exception(httpx.ConnectError("down")) == "retryable_error"


def test_handle_api_error_usage_error_exits_with_code_2(runner, tty_mode):
    @click.command()
    @click.option("--json", "as_json", is_flag=True)
    @common._handle_api_error
    def cmd(as_json: bool):
        raise click.BadParameter("Bad value")

    result = runner.invoke(cmd, ["--json"])

    assert result.exit_code == EXIT_CODE_USAGE
    payload = json.loads(result.stdout)
    assert payload["ok"] is False
    assert payload["error"]["code"] == "invalid_usage"


def test_handle_api_error_retryable_http_error_exits_with_code_8(runner, tty_mode):
    @click.command()
    @click.option("--json", "as_json", is_flag=True)
    @common._handle_api_error
    def cmd(as_json: bool):
        raise httpx.ReadTimeout("timeout")

    result = runner.invoke(cmd, ["--json"])

    assert result.exit_code == EXIT_CODE_RETRYABLE
    payload = json.loads(result.stdout)
    assert payload["ok"] is False
    assert payload["error"]["code"] == "retryable_error"


def test_successful_command_keeps_exit_code_zero(runner, tty_mode):
    @click.command()
    @common._handle_api_error
    def cmd():
        click.echo("ok")

    result = runner.invoke(cmd, [])

    assert result.exit_code == 0
    assert result.output == "ok\n"
