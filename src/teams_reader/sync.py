"""A local archive of chat messages, and notifications for the ones that need you.

``sync`` lists chats ordered by their latest message and fetches only the
chats that changed since the previous run. New messages are appended to one
JSON Lines file per chat under ``$XDG_STATE_HOME/teams-reader``. Run it on a
timer and ``recent`` can answer "what's new" from disk, without calling Graph.

After the first run, ``sync`` can publish an ntfy notification for direct
messages, @mentions of you, and the senders and chats you pick in
``projects.toml``. ``wait`` polls chats live until someone else posts, so an
agent can block on a reply instead of being told about it.

``projects.toml`` (next to ``account.json``) groups chats by project::

    [notify]
    ntfy = "http://127.0.0.1:8090/teams"   # server URL and topic; omit to disable
    people = ["Priya Shah"]                # always notify for these senders
    direct = true                          # direct messages (default true)
    mentions = true                        # @mentions of you (default true)

    [projects.launch]
    chats = ["19:3f2a...@thread.v2"]
    notify = true                          # notify for every message in these chats

Keys teams-reader doesn't use (repos, people, notes) are ignored, so the file
can double as a project registry for agents.
"""
from __future__ import annotations

import fcntl
import json
import os
import re
import sys
import time
import tomllib
import urllib.request
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterator
from urllib.parse import quote, unquote, urlsplit

from . import m365, reader, render
from .timeutil import item_time, parse_since, parse_time

DEFAULT_BACKFILL = '7d'
OVERLAP = timedelta(minutes=10)  # re-list chats this far back, in case Graph lags
STALE_AFTER = timedelta(minutes=15)
WAIT_RETRIES = 5  # consecutive failed polls before `wait` gives up
NOTICE_LINES = 3


def state_dir() -> Path:
    xdg = os.environ.get('XDG_STATE_HOME') or str(Path.home() / '.local/state')
    return Path(xdg) / 'teams-reader'


def projects_path() -> Path:
    return m365.config_paths()[0].with_name('projects.toml')


def load_config() -> dict[str, Any]:
    path = projects_path()
    if not path.exists():
        return {}
    try:
        return tomllib.loads(path.read_text())
    except (OSError, tomllib.TOMLDecodeError) as exc:
        raise m365.ConfigError(f'cannot read {path}: {exc}') from exc


def project_chats(config: dict[str, Any], name: str) -> list[str]:
    projects = config.get('projects') or {}
    if name not in projects:
        known = ', '.join(sorted(projects)) or 'none'
        raise m365.ConfigError(f'unknown project {name!r} in {projects_path()} (known: {known})')
    chats = [chat_id(c) for c in projects[name].get('chats') or []]
    if not chats:
        raise m365.ConfigError(f'project {name!r} has no chats in {projects_path()}')
    return chats


class LinkError(ValueError):
    """A Teams link that names no chat."""


def chat_id(value: str) -> str:
    """A chat ID, also accepting a Teams link to the chat or one of its messages."""
    if not value.startswith(('http://', 'https://')):
        return value
    for segment in unquote(urlsplit(value).path).split('/'):
        if segment.startswith('19:'):
            return segment
    raise LinkError(f'no chat ID in Teams link {value!r}')


def message_link(chat: str, message_id: str) -> str:
    context = quote('{"contextType":"chat"}', safe='')
    return (f"https://teams.microsoft.com/l/message/{quote(chat, safe=':@')}/{message_id}"
            f'?context={context}')


# Archive files and state

def _chat_file(root: Path, chat: str) -> Path:
    return root / 'chats' / (re.sub(r'[^A-Za-z0-9._-]', '_', chat) + '.jsonl')


def _read_archive(path: Path) -> Iterator[dict]:
    try:
        lines = path.read_text().splitlines()
    except FileNotFoundError:
        return
    for line in lines:
        try:
            yield json.loads(line)
        except json.JSONDecodeError:
            continue  # a line still being written by a concurrent sync


def _load_state(root: Path) -> dict[str, Any]:
    try:
        return json.loads((root / 'state.json').read_text())
    except FileNotFoundError:
        return {}


def _save_state(root: Path, state: dict[str, Any]) -> None:
    temp = root / 'state.json.tmp'
    temp.write_text(json.dumps(state, indent=1, ensure_ascii=False) + '\n')
    temp.replace(root / 'state.json')


@contextmanager
def _locked(root: Path) -> Iterator[None]:
    with open(root / '.lock', 'w') as lock:
        fcntl.flock(lock, fcntl.LOCK_EX)
        yield


def _append(root: Path, chat: str, messages: list[dict]) -> list[dict]:
    path = _chat_file(root, chat)
    seen = {m.get('id') for m in _read_archive(path)}
    added = sorted((m for m in messages if m.get('id') not in seen),
                   key=lambda m: m.get('createdDateTime') or '')
    if added:
        with open(path, 'a') as f:
            for message in added:
                f.write(json.dumps({**message, 'chatId': chat}, ensure_ascii=False) + '\n')
    return added


def _me() -> dict[str, str]:
    me = m365.graph_get('/me', {'$select': 'id,displayName'})
    return {'id': me.get('id'), 'name': me.get('displayName')}


def _iso(when: datetime) -> str:
    return when.strftime('%Y-%m-%dT%H:%M:%S.%fZ')


# sync

def sync(backfill: datetime | None = None, notify: bool = True) -> dict[str, Any]:
    """Archive new messages from every chat that changed since the last run."""
    os.umask(0o077)
    config = load_config()
    root = state_dir()
    (root / 'chats').mkdir(parents=True, exist_ok=True)
    with _locked(root):
        state = _load_state(root)
        first_run = 'synced_at' not in state
        started = datetime.now(timezone.utc)
        try:
            state['me'] = state.get('me') or _me()
            new = _fetch_changed(root, state, backfill)
        except m365.M365Error as exc:
            if notify and not state.get('failing'):
                _publish_safely(config, 'Teams sync is failing', str(exc)[:300])
            state['failing'] = str(exc)
            _save_state(root, state)
            raise
        state.pop('failing', None)
        state['synced_at'] = _iso(started)
        _save_state(root, state)
    notices = 0 if first_run or not notify else _notify(config, state, new)
    return {'first_run': first_run, 'chats': len(new),
            'messages': sum(len(m) for m in new.values()), 'notifications': notices}


def _fetch_changed(root: Path, state: dict[str, Any],
                   backfill: datetime | None) -> dict[str, list[dict]]:
    if 'synced_at' in state:
        since = parse_time(state['synced_at']) - OVERLAP
    else:
        since = backfill or parse_since(DEFAULT_BACKFILL)
    known = state.setdefault('chats', {})
    new: dict[str, list[dict]] = {}
    for chat in reader.list_chats(since=since, limit=None):
        cid = chat['id']
        last = parse_time(known[cid]['last']) if cid in known else None
        preview = item_time(chat.get('lastMessagePreview'))
        if last and preview and preview <= last:
            continue
        messages = reader.chat_messages(cid, since=last or since, limit=None,
                                        include_system=False)
        added = _append(root, cid, messages)
        times = [t for t in (last, preview, *(item_time(m) for m in added)) if t]
        known[cid] = {'name': render.chat_name(chat), 'type': chat.get('chatType'),
                      'last': _iso(max(times)) if times else None}
        if added:
            new[cid] = added
    return new


# Notifications

def _sender_id(message: dict) -> str | None:
    return ((message.get('from') or {}).get('user') or {}).get('id')


def _mentions(message: dict, user_id: str | None) -> bool:
    return bool(user_id) and any(
        ((m.get('mentioned') or {}).get('user') or {}).get('id') == user_id
        for m in message.get('mentions') or [])


def _wants_notice(message: dict, chat: str, chat_type: str | None, me: str | None,
                  settings: dict[str, Any], people: set[str], loud: set[str]) -> bool:
    if me and _sender_id(message) == me:
        return False
    return (chat in loud or render.sender(message).casefold() in people
            or (settings.get('direct', True) and chat_type == 'oneOnOne')
            or (settings.get('mentions', True) and _mentions(message, me)))


def _notify(config: dict[str, Any], state: dict[str, Any],
            new: dict[str, list[dict]]) -> int:
    settings = config.get('notify') or {}
    if not settings.get('ntfy') or not new:
        return 0
    me = (state.get('me') or {}).get('id')
    people = {p.casefold() for p in settings.get('people') or []}
    loud = {chat_id(c) for p in (config.get('projects') or {}).values() if p.get('notify')
            for c in p.get('chats') or []}
    sent = 0
    for cid, messages in new.items():
        info = state['chats'][cid]
        wanted = [m for m in messages
                  if _wants_notice(m, cid, info.get('type'), me, settings, people, loud)]
        if not wanted:
            continue
        senders = list(dict.fromkeys(render.sender(m) for m in wanted))
        where = '' if info.get('type') == 'oneOnOne' else f" · {info.get('name')}"
        title = (senders[0] if len(senders) == 1 else f'{len(wanted)} new') + where
        lines = [f'{render.sender(m)}: {render._clip(render.message_text(m, quotes=False), 200)}'
                 for m in wanted[-NOTICE_LINES:]]
        if len(wanted) > NOTICE_LINES:
            lines.insert(0, f'(+{len(wanted) - NOTICE_LINES} earlier)')
        if _publish_safely(config, title, '\n'.join(lines), message_link(cid, wanted[-1]['id'])):
            sent += 1
    return sent


def _publish_safely(config: dict[str, Any], title: str, message: str,
                    click: str | None = None) -> bool:
    url = ((config.get('notify') or {}).get('ntfy') or '').rstrip('/')
    if not url:
        return False
    server, _, topic = url.rpartition('/')
    body: dict[str, str] = {'topic': topic, 'title': title, 'message': message}
    if click:
        body['click'] = click
    request = urllib.request.Request(server, data=json.dumps(body).encode(),
                                     headers={'Content-Type': 'application/json'})
    try:
        urllib.request.urlopen(request, timeout=10).close()
    except OSError as exc:
        print(f'teams-reader: warning: ntfy notification failed: {exc}', file=sys.stderr)
        return False
    return True


# recent

def recent(chats: list[str] | None = None, since: datetime | None = None,
           limit: int | None = reader.DEFAULT_LIMIT, sender: str | None = None
           ) -> tuple[list[dict], list[str]]:
    """Archived messages, newest first, each with ``chatName``. Also returns warnings."""
    root = state_dir()
    state = _load_state(root)
    if 'synced_at' not in state:
        raise m365.ConfigError('no local archive yet: run `teams-reader sync` first')
    warnings = []
    age = datetime.now(timezone.utc) - parse_time(state['synced_at'])
    if age > STALE_AFTER:
        warnings.append(f'the archive was last synced {int(age.total_seconds() // 60)} minutes ago')
    if state.get('failing'):
        warnings.append(f"the last sync failed: {state['failing']}")
    files = ([_chat_file(root, c) for c in chats] if chats
             else sorted((root / 'chats').glob('*.jsonl')))
    names = {cid: info.get('name') for cid, info in (state.get('chats') or {}).items()}
    wanted = sender.casefold() if sender else None
    items = []
    for path in files:
        for message in _read_archive(path):
            created = item_time(message)
            if since and (created is None or created < since):
                continue
            if wanted and wanted not in render.sender(message).casefold():
                continue
            items.append({**message, 'chatName': names.get(message.get('chatId'))})
    items.sort(key=lambda m: m.get('createdDateTime') or '', reverse=True)
    return (items[:limit] if limit else items), warnings


# wait

def wait(chats: list[str], since: datetime | None = None, sender: str | None = None,
         timeout: timedelta = timedelta(hours=8), interval: float = 60) -> list[dict]:
    """Poll chats until someone else posts after ``since`` (default now).

    Returns the new messages, newest first, or an empty list on timeout.
    """
    since = since or datetime.now(timezone.utc)
    me = _me().get('id')
    wanted = sender.casefold() if sender else None
    deadline = time.monotonic() + timeout.total_seconds()
    failures = 0
    while True:
        try:
            found = [{**m, 'chatId': cid} for cid in chats
                     for m in reader.chat_messages(cid, since=since, limit=None,
                                                   include_system=False)
                     if _sender_id(m) != me
                     and (not wanted or wanted in render.sender(m).casefold())]
            failures = 0
        except m365.M365Error:
            failures += 1
            if failures >= WAIT_RETRIES:
                raise
            found = []
        if found:
            return sorted(found, key=lambda m: m.get('createdDateTime') or '', reverse=True)
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return []
        time.sleep(min(interval, remaining))
