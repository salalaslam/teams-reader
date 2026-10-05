"""Searching Teams messages, two ways.

By default, Microsoft Search (``m365 search --scopes chatMessage``) queries the
user's whole chat and channel history server-side. It matches words and
prefixes using KQL syntax and returns a highlighted snippet per message.

With ``scan=True`` the search runs client-side instead: recent messages are
read from the most recently active chats and from every channel of every
joined team, and matched as case-insensitive substrings of the full text.
Use it for exact fragments (IDs, partial words) or when Microsoft Search is
unavailable.
"""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from functools import partial
from typing import Callable

from . import m365, reader, render
from .timeutil import item_time, parse_since

Job = tuple[str, Callable[[], list]]  # (label for warnings, work)

DEFAULT_LIMIT = 25
DEFAULT_SCAN_SINCE = '7d'
DEFAULT_MAX_CHATS = 25
PER_SOURCE_LIMIT = 200  # messages scanned per chat or channel
WORKERS = 4  # concurrent m365 processes


def search(query: str, since: datetime | None = None, *, limit: int | None = DEFAULT_LIMIT,
           channels: bool = True, scan: bool = False, max_chats: int = DEFAULT_MAX_CHATS
           ) -> tuple[list[dict], list[str]]:
    """Return (matches newest first, warnings about sources that failed)."""
    if not query.strip():
        raise ValueError('search query must not be empty')
    if scan:
        return _scan(query, since or parse_since(DEFAULT_SCAN_SINCE), limit=limit,
                     channels=channels, max_chats=max_chats)
    return _microsoft_search(query, since, limit=limit, channels=channels)


def _microsoft_search(query: str, since: datetime | None, *, limit: int | None,
                      channels: bool) -> tuple[list[dict], list[str]]:
    kql = query.strip()
    if since is not None:
        # KQL dates are whole days; narrow to the exact time below.
        kql += f' sent>={(since - timedelta(days=1)).date().isoformat()}'
    size = min(limit * 2, m365.SEARCH_PAGE_MAX) if limit else m365.SEARCH_PAGE_MAX
    matches = []
    for hit in m365.search_messages(kql, size):
        record = _search_hit(hit)
        created = item_time(record)
        if since is not None and (created is None or created < since):
            continue
        if record['source'] == 'channel' and not channels:
            continue
        matches.append(record)
    matches.sort(key=lambda m: m['createdDateTime'] or '', reverse=True)
    matches = matches[:limit] if limit else matches
    warnings: list[str] = []
    _add_names(matches, warnings)
    return matches, warnings


def _search_hit(hit: dict) -> dict:
    resource = hit.get('resource') or {}
    identity = resource.get('channelIdentity') or {}
    if identity.get('channelId'):
        where = dict(source='channel', teamId=identity.get('teamId'), teamName=None,
                     channelId=identity.get('channelId'), channelName=None)
    else:
        where = dict(source='chat', chatId=resource.get('chatId'), chatName=None)
    sender = ((resource.get('from') or {}).get('emailAddress') or {}).get('name')
    return {
        **where,
        'messageId': resource.get('id'),
        'createdDateTime': resource.get('createdDateTime'),
        'from': sender or 'unknown',
        'text': render.html_to_text(hit.get('summary')),
        'webUrl': resource.get('webLink'),
    }


def _add_names(matches: list[dict], warnings: list[str]) -> None:
    """Fill in chat, team and channel names, which search results lack."""
    chat_ids = sorted({m['chatId'] for m in matches if m['source'] == 'chat'})
    channel_ids = sorted({(m['teamId'], m['channelId']) for m in matches
                          if m['source'] == 'channel'})
    jobs: list[Job] = [(f'chat {c}', partial(_one, reader.get_chat, c)) for c in chat_ids]
    jobs += [(f'channel {c}', partial(_one, reader.get_channel, t, c)) for t, c in channel_ids]
    if channel_ids:
        jobs.append(('teams list', reader.list_teams))
    results = _parallel(jobs, warnings)
    chats = {c: render.chat_name(r[0]) for c, r in zip(chat_ids, results) if r}
    names = {key: r[0].get('displayName') for key, r in
             zip(channel_ids, results[len(chat_ids):]) if r}
    teams = {t['id']: t.get('displayName') for t in (results[-1] if channel_ids else [])}
    for m in matches:
        if m['source'] == 'chat':
            m['chatName'] = chats.get(m['chatId'])
        else:
            m['teamName'] = teams.get(m['teamId'])
            m['channelName'] = names.get((m['teamId'], m['channelId']))


def _one(fetch: Callable[..., dict], *ids: str) -> list[dict]:
    return [fetch(*ids)]


def _scan(query: str, since: datetime, *, limit: int | None, max_chats: int,
          channels: bool) -> tuple[list[dict], list[str]]:
    needle = query.casefold().strip()
    warnings: list[str] = []
    jobs: list[Job] = []

    # Sequential calls first, so any token refresh happens before fan-out.
    for chat in reader.list_chats(since=since, limit=max_chats):
        name = render.chat_name(chat)
        jobs.append((f'chat {name}', partial(_scan_chat, chat['id'], name, since)))
    if channels:
        teams = reader.list_teams()
        listings = _parallel([(f"team {t.get('displayName')}",
                               partial(reader.list_channels, t['id'])) for t in teams], warnings)
        for team, team_channels in zip(teams, listings):
            for channel in team_channels:
                jobs.append((f"channel {team.get('displayName')} / {channel.get('displayName')}",
                             partial(_scan_channel, team, channel, since)))

    matches = [hit for batch in _parallel(jobs, warnings) for hit, own_text in batch
               if needle in own_text.casefold()]
    matches.sort(key=lambda m: m['createdDateTime'] or '', reverse=True)
    return (matches[:limit] if limit else matches), warnings


def _scan_chat(chat_id: str, chat_name: str, since: datetime) -> list[tuple[dict, str]]:
    found = reader.chat_messages(chat_id, since=since, limit=PER_SOURCE_LIMIT,
                                 include_system=False)
    return [_hit(m, source='chat', chatId=chat_id, chatName=chat_name) for m in found]


def _scan_channel(team: dict, channel: dict, since: datetime) -> list[tuple[dict, str]]:
    where = dict(source='channel', teamId=team['id'], teamName=team.get('displayName'),
                 channelId=channel['id'], channelName=channel.get('displayName'))
    posts = reader.channel_posts(team['id'], channel['id'], since=since,
                                 limit=PER_SOURCE_LIMIT, include_system=False, replies=True)
    hits = []
    for post in posts:
        for message in [post, *post.get('replies', [])]:
            created = item_time(message)
            if created is not None and created >= since:
                hits.append(_hit(message, threadId=post['id'], **where))
    return hits


def _hit(message: dict, **where) -> tuple[dict, str]:
    """A result record, plus the text to match: quoted messages are excluded
    so a reply doesn't match on the message it quotes."""
    return {
        **where,
        'messageId': message.get('id'),
        'createdDateTime': message.get('createdDateTime'),
        'from': render.sender(message),
        'text': render.message_text(message),
        'webUrl': message.get('webUrl'),
    }, render.message_text(message, quotes=False)


def _parallel(jobs: list[Job], warnings: list[str]) -> list[list]:
    """Run jobs concurrently; a failed source becomes a warning and an empty list."""

    def run(job: Job) -> list:
        label, work = job
        try:
            return work()
        except m365.M365Error as exc:
            warnings.append(f'skipped {label}: {exc}')
            return []

    if not jobs:
        return []
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        return list(pool.map(run, jobs))
