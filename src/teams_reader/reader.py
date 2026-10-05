"""Read-only Teams operations, returning Microsoft Graph objects."""
from __future__ import annotations

from datetime import datetime

from . import m365
from .timeutil import item_time

DEFAULT_LIMIT = 50


def is_system_event(message: dict) -> bool:
    """True for membership/call/rename events and deleted messages."""
    return message.get('messageType') != 'message' or bool(message.get('deletedDateTime'))


def _filters(since: datetime | None, include_system: bool):
    def too_old(item: dict) -> bool:
        created = item_time(item)
        return since is not None and created is not None and created < since

    def keep(item: dict) -> bool:
        return (include_system or not is_system_event(item)) and not too_old(item)

    return keep, too_old


def list_chats(since: datetime | None = None, limit: int | None = DEFAULT_LIMIT) -> list[dict]:
    """Chats ordered by most recent message, with members and a preview."""

    def last_activity(chat: dict):
        return item_time(chat.get('lastMessagePreview'))

    def keep(chat: dict) -> bool:
        if since is None:
            return True
        last = last_activity(chat)
        return last is not None and last >= since

    def stop(chat: dict) -> bool:
        last = last_activity(chat)
        return since is not None and last is not None and last < since

    return m365.paginate('/me/chats', {
        '$top': m365.GRAPH_PAGE_MAX,
        '$expand': 'members,lastMessagePreview',
        '$orderby': 'lastMessagePreview/createdDateTime desc',
    }, limit=limit, keep=keep, stop=stop)


def chat_messages(chat_id: str, since: datetime | None = None,
                  limit: int | None = DEFAULT_LIMIT,
                  include_system: bool = True) -> list[dict]:
    """Most recent messages in a chat, newest first."""
    keep, too_old = _filters(since, include_system)
    return m365.paginate(m365.path('/chats/{}/messages', chat_id),
                         {'$top': m365.GRAPH_PAGE_MAX, '$orderby': 'createdDateTime desc'},
                         limit=limit, keep=keep, stop=too_old)


def list_teams() -> list[dict]:
    return m365.paginate('/me/joinedTeams')


def list_channels(team_id: str) -> list[dict]:
    return m365.paginate(m365.path('/teams/{}/channels', team_id))


def channel_posts(team_id: str, channel_id: str, since: datetime | None = None,
                  limit: int | None = DEFAULT_LIMIT, include_system: bool = True,
                  replies: bool = False) -> list[dict]:
    """Most recent root posts in a channel, newest first.

    Graph offers no server-side ordering or date filter here, so posts are
    filtered client-side and paging stops at the first post older than
    ``since`` (Graph returns channel posts newest first).
    """
    keep, too_old = _filters(since, include_system)
    params = {'$top': m365.GRAPH_PAGE_MAX}
    if replies:
        params['$expand'] = 'replies'
    return m365.paginate(m365.path('/teams/{}/channels/{}/messages', team_id, channel_id),
                         params, limit=limit, keep=keep, stop=too_old)


def post_replies(team_id: str, channel_id: str, message_id: str,
                 since: datetime | None = None, limit: int | None = DEFAULT_LIMIT,
                 include_system: bool = True) -> list[dict]:
    """Replies to a channel post. Filtering is client-side."""
    keep, _ = _filters(since, include_system)
    return m365.paginate(
        m365.path('/teams/{}/channels/{}/messages/{}/replies', team_id, channel_id, message_id),
        {'$top': m365.GRAPH_PAGE_MAX}, limit=limit, keep=keep)
