"""Stable per-item mutation outcomes without dropping earlier successes."""
from concurrent.futures import ThreadPoolExecutor

import click

from ..exceptions import error_code_for_exception
from ._common import _wants_json, to_json_envelope


def run_items(ids, operation, action, *, workers=1):
    def invoke(item_id):
        try:
            data = action(item_id)
            return {"id": item_id, "ok": True, "operation": operation, "data": data}
        except Exception as exc:
            return {"id": item_id, "ok": False, "operation": operation, "error": {"code": error_code_for_exception(exc), "message": str(exc)}}
    with ThreadPoolExecutor(max_workers=min(workers, max(1, len(ids)))) as pool:
        results = list(pool.map(invoke, ids))
    failed = sum(not result["ok"] for result in results)
    if _wants_json(False):
        click.echo(to_json_envelope(results, ok=not failed, meta={"returned": len(results), "failed": failed, "partial": bool(failed and failed < len(results))}, error={"code": "partial_failure", "message": f"{failed} item(s) failed"} if failed else None))
    elif failed:
        for item in results:
            if not item["ok"]:
                click.echo(f"{item['id']}: {item['error']['message']}", err=True)
    if failed:
        raise click.exceptions.Exit(1)
    return results
