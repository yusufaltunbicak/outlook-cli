"""Shared recipient parsing: repeated options, commas and semicolons."""
from __future__ import annotations

import re

import click

_ADDRESS = re.compile(r"^[^\s<>@,;]+@[^\s<>@,;]+\.[^\s<>@,;]+$")


def normalize_recipients(values, *, field="TO", required=False) -> list[str]:
    if isinstance(values, str):
        values = [values]
    chunks = list(values or [])
    if any("\n" in part or "\r" in part for part in chunks):
        raise click.BadParameter(f"{field} cannot contain newlines", param_hint=field)
    result = []
    seen = set()
    for raw in chunks:
        # Split separators outside quoted display names / angle addresses.
        parts, start, quoted, angle = [], 0, False, 0
        for index, char in enumerate(raw):
            if char == '"' and (index == 0 or raw[index - 1] != "\\"):
                quoted = not quoted
            elif not quoted and char == "<":
                angle += 1
            elif not quoted and char == ">":
                angle -= 1
            elif not quoted and not angle and char in ",;":
                parts.append(raw[start:index])
                start = index + 1
        parts.append(raw[start:])
        if quoted or angle:
            raise click.BadParameter(f"Invalid {field} recipient: {raw!r}", param_hint=field)
        for part in parts:
            part = part.strip()
            if not part:
                continue
            if "<" in part:
                match = re.fullmatch(r'[^<>]*<([^<>]+)>', part)
                address = match.group(1).strip() if match else ""
            else:
                address = part
            if not _ADDRESS.fullmatch(address):
                raise click.BadParameter(f"Invalid {field} email address: {address or raw!r}", param_hint=field)
            key = address.casefold()
            if key not in seen:
                seen.add(key)
                result.append(address)
    if required and not result:
        raise click.BadParameter(f"At least one {field} email address is required", param_hint=field)
    return result
