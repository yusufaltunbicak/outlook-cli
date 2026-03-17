"""Manage email signatures for outlook-cli.

Signatures are stored as HTML files in ~/.config/outlook-cli/signatures/.
They can be pulled from sent emails or created manually.
"""

from __future__ import annotations

from pathlib import Path

import httpx

from . import account as account_service
from .constants import BASE_URL
from .exceptions import ResourceNotFoundError


def _signatures_dir(account_name: str | None = None) -> Path:
    selected = account_service.resolve_account_name(account_name)
    return account_service.get_account_paths(selected).signatures_dir


def list_signatures(account_name: str | None = None) -> list[str]:
    """Return names of saved signatures (without .html extension)."""
    signatures_dir = _signatures_dir(account_name)
    if not signatures_dir.exists():
        return []
    return sorted(p.stem for p in signatures_dir.glob("*.html"))


def get_signature(name: str, account_name: str | None = None) -> str:
    """Load a signature's HTML content by name."""
    path = _signatures_dir(account_name) / f"{name}.html"
    if not path.exists():
        raise ResourceNotFoundError(
            f"Signature '{name}' not found. Run 'outlook signature list' to see available signatures."
        )
    return path.read_text(encoding="utf-8")


def save_signature(name: str, html: str, account_name: str | None = None) -> Path:
    """Save signature HTML to disk."""
    signatures_dir = _signatures_dir(account_name)
    signatures_dir.mkdir(parents=True, exist_ok=True)
    path = signatures_dir / f"{name}.html"
    path.write_text(html, encoding="utf-8")
    return path


def delete_signature(name: str, account_name: str | None = None) -> None:
    """Delete a saved signature."""
    path = _signatures_dir(account_name) / f"{name}.html"
    if not path.exists():
        raise ResourceNotFoundError(f"Signature '{name}' not found.")
    path.unlink()


def pull_signature(token: str) -> tuple[str, str]:
    """Extract signature HTML from the most recent sent email that contains one.

    Looks for emails with structured signature blocks (containing phone/email/company info).

    Returns (signature_html, source_subject).
    """
    client = httpx.Client(
        base_url=BASE_URL,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        timeout=30,
    )
    try:
        for skip in range(0, 50, 5):
            resp = client.get(
                "/MailFolders/SentItems/messages",
                params={
                    "$top": 5,
                    "$skip": skip,
                    "$orderby": "SentDateTime desc",
                    "$select": "Subject,Body",
                },
            )
            if resp.status_code != 200:
                break
            for m in resp.json().get("value", []):
                body = m.get("Body", {}).get("Content", "")
                subject = m.get("Subject", "")
                sig = _extract_signature(body)
                if sig:
                    return sig, subject
    finally:
        client.close()
    raise ResourceNotFoundError(
        "Could not find a signature in your sent emails. "
        "Send an email with your signature from Outlook first."
    )


def _extract_signature(html: str) -> str | None:
    """Extract the signature table block from an email HTML body.

    Strategy: find all top-level <table> blocks before the reply chain,
    then pick the one containing mailto: links (signature indicator).
    The outermost table containing that block is the full signature.
    """
    # Cut content before reply chain
    reply_markers = ["divRplyFwdMsg", "x_divRplyFwdMsg", "gmail_quote", "originalMessage"]
    main_body = html
    for marker in reply_markers:
        idx = html.find(marker)
        if idx > 0:
            # Go back to find the enclosing div/hr
            cut = html.rfind("<hr", 0, idx)
            if cut < 0:
                cut = html.rfind("<div", 0, idx)
            main_body = html[:cut] if cut > 0 else html[:idx]
            break

    if "mailto:" not in main_body:
        return None

    # Find all <table starts in the main body
    table_starts = []
    pos = 0
    while True:
        idx = main_body.find("<table", pos)
        if idx < 0:
            break
        table_starts.append(idx)
        pos = idx + 1

    if not table_starts:
        return None

    # For each table start, extract the full table and check if it contains mailto:
    # Go from earliest to find the outermost signature container
    best = None
    for start in table_starts:
        # Extract balanced table
        table_html = _extract_balanced_table(main_body, start)
        if table_html and "mailto:" in table_html:
            # Prefer the outermost (earliest) table that contains mailto
            if best is None or start < best[0]:
                best = (start, table_html)

    if best is None:
        return None

    sig_html = best[1]

    # Validate: must be substantial
    if len(sig_html) < 100:
        return None

    return sig_html


def _extract_balanced_table(html: str, start: int) -> str | None:
    """Extract a complete <table>...</table> block with balanced nesting."""
    depth = 0
    i = start
    while i < len(html):
        if html[i:].startswith("<table"):
            depth += 1
            i += 6
        elif html[i:].startswith("</table>"):
            depth -= 1
            i += 8
            if depth == 0:
                return html[start:i]
        else:
            i += 1
    return None


def append_signature(body: str, signature_html: str, is_html: bool) -> tuple[str, bool]:
    """Append signature to email body.

    If body is plain text and signature is HTML, wraps body in HTML.
    Returns (new_body, is_html).
    """
    sig_block = f'<br><div id="Signature">{signature_html}<br></div>'

    if is_html:
        # Insert signature before closing </body> or </html> or at the end
        for tag in ["</body>", "</html>"]:
            idx = body.lower().rfind(tag)
            if idx > 0:
                return body[:idx] + sig_block + body[idx:], True
        return body + sig_block, True
    else:
        # Plain text body + HTML signature → convert to HTML
        escaped = body.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        html_body = escaped.replace("\n", "<br>")
        full = (
            '<html><body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000;">'
            f"<div>{html_body}</div>"
            f"{sig_block}"
            "</body></html>"
        )
        return full, True
