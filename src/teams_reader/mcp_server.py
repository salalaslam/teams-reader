"""Stdio MCP server exposing the same read-only operations as the CLI.

Requires the optional ``mcp`` extra: ``pip install 'teams-reader[mcp]'``.
Every tool is a read: there are no tools for login, logout, sending,
editing or deleting.
"""
from __future__ import annotations

import functools
import json
from typing import Any, Literal

from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations

from . import __version__, m365, reader, render, search
from .timeutil import parse_since

Format = Literal['md', 'json']

READ_ONLY = ToolAnnotations(read_only_hint=True, destructive_hint=False,
                            idempotent_hint=True, open_world_hint=True)

server = MCPServer(
    'teams-reader',
    version=__version__,
    instructions=(
        'Read-only access to the signed-in user\'s Microsoft Teams chats and channels. '
        'Start with list_chats or search_messages; use IDs from results in follow-up '
        'calls. since accepts 7d, 12h, 30m, 2w or an ISO date. Times are UTC. '
        'format="md" (default) is compact text; use "json" for full Graph objects.'),
)


def tool(fn):
    """Register a read-only tool; expected failures reach the model as messages."""

    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except (m365.M365Error, m365.ConfigError, ValueError) as exc:
            raise ToolError(str(exc)) from exc

    return server.tool(annotations=READ_ONLY, structured_output=False)(wrapper)


def _out(kind: str, data: Any, fmt: Format) -> str:
    if fmt == 'json':
        return json.dumps(data, ensure_ascii=False)
    return render.RENDERERS[kind](data)


def _limit(limit: int) -> int | None:
    if limit < 0:
        raise ValueError('limit must be >= 0')
    return limit or None


@tool
def teams_status() -> str:
    """Show which Microsoft 365 account and Entra app the reader is signed in with."""
    return _out('status', m365.status(), 'md')


@tool
def list_chats(since: str | None = None, limit: int = 30, format: Format = 'md') -> str:
    """List the user's chats (1:1, group, meeting), most recently active first,
    with members and a preview of the last message.

    since: only chats with a message at or after this time (e.g. 7d, 2026-10-01).
    limit: maximum chats (0 = no limit).
    """
    return _out('chats', reader.list_chats(since=parse_since(since), limit=_limit(limit)),
                format)


@tool
def read_chat(chat_id: str, since: str | None = None, limit: int = 50,
              format: Format = 'md') -> str:
    """Read the most recent messages in a chat. Markdown output is oldest first and
    omits system events; JSON is newest first and includes everything.

    since: only messages sent at or after this time. limit: maximum messages (0 = all).
    """
    return _out('messages', reader.chat_messages(
        chat_id, since=parse_since(since), limit=_limit(limit),
        include_system=format == 'json'), format)


@tool
def list_teams(format: Format = 'md') -> str:
    """List the teams the user is a member of."""
    return _out('teams', reader.list_teams(), format)


@tool
def list_channels(team_id: str, format: Format = 'md') -> str:
    """List the channels the user can see in a team."""
    return _out('channels', reader.list_channels(team_id), format)


@tool
def read_channel(team_id: str, channel_id: str, since: str | None = None, limit: int = 20,
                 include_replies: bool = True, format: Format = 'md') -> str:
    """Read the most recently active threads in a channel, optionally with replies.

    since: only threads with a post or reply at or after this time.
    limit: maximum threads (0 = no limit).
    """
    return _out('posts', reader.channel_posts(
        team_id, channel_id, since=parse_since(since), limit=_limit(limit),
        include_system=format == 'json', replies=include_replies), format)


@tool
def read_thread_replies(team_id: str, channel_id: str, message_id: str,
                        since: str | None = None, limit: int = 50,
                        format: Format = 'md') -> str:
    """Read replies to one channel post (message_id is the thread's root post ID)."""
    return _out('replies', reader.post_replies(
        team_id, channel_id, message_id, since=parse_since(since), limit=_limit(limit),
        include_system=format == 'json'), format)


@tool
def search_messages(query: str, since: str = search.DEFAULT_SINCE, limit: int = 30,
                    max_chats: int = search.DEFAULT_MAX_CHATS, include_channels: bool = True,
                    format: Format = 'md') -> str:
    """Find messages containing `query` (case-insensitive substring) in recently
    active chats and in all channels of the user's teams. This scans messages
    client-side, so keep `since` short (default 7d); it can take 10-60 seconds.

    max_chats: scan at most this many of the most recently active chats.
    JSON output is {"matches": [...], "warnings": [...]}.
    """
    since_time = parse_since(since)
    if since_time is None:
        raise ValueError('since is required for search')
    matches, warnings = search.search(query, since_time, limit=_limit(limit),
                                      max_chats=max_chats, channels=include_channels)
    if format == 'json':
        return json.dumps({'matches': matches, 'warnings': warnings}, ensure_ascii=False)
    text = render.search_results(matches)
    if warnings:
        text += '\n\nWarnings:\n' + '\n'.join(f'- {w}' for w in warnings)
    return text


def main() -> None:
    server.run('stdio')
