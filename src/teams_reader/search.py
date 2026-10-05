"""Client-side search across recent chats and channel threads.

Graph's message search (POST /search/query) is not exposed by m365 as a
read-only command, so this scans recent messages instead: the most recently
active chats, plus every channel of every joined team, within a time window.
"""
from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from functools import partial
from typing import Callable

from . import m365, reader, render
from .timeutil import item_time

Job = tuple[str, Callable[[], list]]  # (label for warnings, work)

DEFAULT_SINCE = '7d'
DEFAULT_MAX_CHATS = 25
PER_SOURCE_LIMIT = 200  # messages scanned per chat or channel
WORKERS = 4  # concurrent m365 processes


def search(query: str, since: datetime, *, limit: int | None = 50,
           max_chats: int = DEFAULT_MAX_CHATS, channels: bool = True
           ) -> tuple[list[dict], list[str]]:
    """Return (matches newest first, warnings for sources that failed)."""
    needle = query.casefold().strip()
    if not needle:
        raise ValueError('search query must not be empty')
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
