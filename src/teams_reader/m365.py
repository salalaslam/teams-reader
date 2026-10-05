"""Thin, read-only bridge to the CLI for Microsoft 365 (`m365`).

Only three m365 commands are ever run for reading:

- ``m365 request --method get`` via :func:`graph_get`. No caller can supply an
  arbitrary URL: paths are built from fixed templates with each ID
  percent-encoded, and only Graph nextLinks are followed.
- ``m365 search --scopes chatMessage`` via :func:`search_messages`, m365's
  built-in Microsoft Search query (a read, though Graph takes it as a POST).
- ``m365 status``.
"""
from __future__ import annotations

import json
import os
import shlex
import shutil
import subprocess
from pathlib import Path
from typing import Any, Callable, Iterator
from urllib.parse import quote

GRAPH_ROOT = 'https://graph.microsoft.com/v1.0'
GRAPH_PAGE_MAX = 50  # Graph's maximum $top for the Teams endpoints used here
SEARCH_PAGE_MAX = 500  # m365 search --pageSize maximum



def legacy_root() -> Path:
    """Install location of the original shell-script deployment. Still honoured
    so existing installs keep working without reconfiguration."""
    return Path.home() / '.local/share/teams-reader'


class M365Error(RuntimeError):
    """The m365 CLI failed or is unavailable."""


class ConfigError(RuntimeError):
    """Login configuration (Entra app ID / tenant) is missing or invalid."""


def find_cli() -> list[str]:
    """Return the argv prefix used to invoke m365."""
    override = os.environ.get('TEAMS_READER_M365')
    if override:
        return shlex.split(override)
    on_path = shutil.which('m365')
    if on_path:
        return [on_path]
    legacy = legacy_root() / 'node_modules/.bin/m365'
    if legacy.exists():
        return [str(legacy)]
    raise M365Error(
        'CLI for Microsoft 365 (m365) not found. Install it with '
        '`npm install -g @pnp/cli-microsoft365`, or set TEAMS_READER_M365 '
        'to its path.')


def run(args: list[str]) -> Any:
    """Run an m365 command with JSON output and return the parsed result."""
    os.umask(0o077)  # keep any token-cache files m365 creates private
    argv = [*find_cli(), *args, '--output', 'json']
    try:
        proc = subprocess.run(argv, capture_output=True, text=True)
    except OSError as exc:
        raise M365Error(f'could not run m365: {exc}') from exc
    if proc.returncode != 0:
        message = (proc.stderr or proc.stdout).strip()
        if message.startswith('Error: '):
            message = message[len('Error: '):]
        raise M365Error(message or f'm365 exited with status {proc.returncode}')
    out = proc.stdout.strip()
    if not out:
        return None
    try:
        return json.loads(out)
    except json.JSONDecodeError as exc:
        raise M365Error(f'unexpected m365 output: {out[:200]}') from exc


def path(template: str, *ids: str) -> str:
    """Fill a Graph path template, percent-encoding every ID segment."""
    return template.format(*(quote(i, safe='') for i in ids))


def graph_get(path_or_url: str, params: dict[str, Any] | None = None) -> Any:
    """GET a Graph v1.0 resource. Accepts a path or a Graph nextLink URL."""
    if path_or_url.startswith('https://'):
        if not path_or_url.startswith(GRAPH_ROOT + '/'):
            raise M365Error(f'refusing to follow non-Graph link: {path_or_url}')
        url = path_or_url
    else:
        url = GRAPH_ROOT + path_or_url
    if params:
        query = '&'.join(f'{k}={quote(str(v), safe="/,")}' for k, v in params.items())
        url += ('&' if '?' in url else '?') + query
    return run(['request', '--method', 'get', '--url', url])


def paginate(path: str, params: dict[str, Any] | None = None, *,
             limit: int | None = None,
             keep: Callable[[dict], bool] = lambda item: True,
             stop: Callable[[dict], bool] = lambda item: False) -> list[dict]:
    """Collect items across Graph pages.

    ``keep`` filters items; ``stop`` ends paging at the first item for which it
    is true (used with newest-first ordering and a ``since`` cut-off).
    ``limit`` caps the number of kept items; ``None`` or 0 means no cap. If
    ``params`` has a ``$top`` page size (only for endpoints that support it),
    it is lowered to ``limit`` so small requests fetch a single small page.
    """
    params = dict(params or {})
    if limit and '$top' in params:
        params['$top'] = min(limit, int(params['$top']))
    results: list[dict] = []
    for item in _iter_items(path, params):
        if stop(item):
            break
        if keep(item):
            results.append(item)
            if limit and len(results) >= limit:
                break
    return results


def _iter_items(path: str, params: dict[str, Any]) -> Iterator[dict]:
    page = graph_get(path, params)
    while True:
        yield from (page or {}).get('value', [])
        next_link = (page or {}).get('@odata.nextLink')
        if not next_link:
            return
        page = graph_get(next_link)


def search_messages(query: str, size: int) -> list[dict]:
    """Microsoft Search over chat and channel messages, newest first."""
    hits = run(['search', '--scopes', 'chatMessage', f'--queryText={query}',
                '--pageSize', str(max(1, min(size, SEARCH_PAGE_MAX))), '--resultsOnly'])
    return hits or []


def status() -> Any:
    return run(['status'])


def logout() -> int:
    return subprocess.call([*find_cli(), 'logout'])


def login(app_id: str | None = None, tenant: str | None = None) -> int:
    """Start interactive device-code sign-in with the user's own Entra app."""
    app_id, tenant = load_account(app_id, tenant)
    os.umask(0o077)
    return subprocess.call([
        *find_cli(), 'login', '--authType', 'deviceCode', '--appId', app_id,
        '--tenant', tenant, '--connectionName', 'teams-reader'])


def config_paths() -> list[Path]:
    xdg = os.environ.get('XDG_CONFIG_HOME') or str(Path.home() / '.config')
    return [Path(xdg) / 'teams-reader/account.json', legacy_root() / 'account.json']


def load_account(app_id: str | None = None, tenant: str | None = None) -> tuple[str, str]:
    """Resolve the Entra app ID and tenant: flags, then env, then account.json."""
    app_id = app_id or os.environ.get('TEAMS_READER_APP_ID')
    tenant = tenant or os.environ.get('TEAMS_READER_TENANT')
    if not (app_id and tenant):
        for candidate in config_paths():
            if candidate.exists():
                try:
                    data = json.loads(candidate.read_text())
                except (OSError, json.JSONDecodeError) as exc:
                    raise ConfigError(f'cannot read {candidate}: {exc}') from exc
                app_id = app_id or data.get('appId')
                tenant = tenant or data.get('tenant')
                break
    if not (app_id and tenant):
        raise ConfigError(
            'No Entra app configured. Pass --app-id and --tenant, set '
            'TEAMS_READER_APP_ID and TEAMS_READER_TENANT, or create '
            f'{config_paths()[0]} with "appId" and "tenant".')
    return app_id, tenant
