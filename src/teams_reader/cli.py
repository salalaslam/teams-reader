"""Read-only Microsoft Teams CLI for AI agents, built on CLI for Microsoft 365."""
from __future__ import annotations

import argparse
import json
import sys

from . import __version__, m365, reader
from .timeutil import parse_since


def _count(value: str) -> int:
    try:
        number = int(value)
    except ValueError:
        number = -1
    if number < 0:
        raise argparse.ArgumentTypeError(f'expected a whole number >= 0, got {value!r}')
    return number


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog='teams-reader', description=__doc__,
        epilog='Times are UTC. --since accepts 7d, 12h, 30m, 2w, or an ISO date/date-time.')
    parser.add_argument('--version', action='version', version=f'%(prog)s {__version__}')
    sub = parser.add_subparsers(dest='command', required=True, metavar='COMMAND')

    def window(p: argparse.ArgumentParser, default_limit: int = reader.DEFAULT_LIMIT) -> None:
        p.add_argument('--since', metavar='WHEN',
                       help='only items created at or after WHEN (e.g. 7d, 2026-10-01)')
        p.add_argument('--limit', '--top', '-n', type=_count, default=default_limit, metavar='N',
                       help=f'maximum items to return (default {default_limit}; 0 = no limit)')

    sub.add_parser('status', help='show the signed-in account')
    login = sub.add_parser('login', help='sign in interactively with a device code')
    login.add_argument('--app-id', help='Entra application (client) ID')
    login.add_argument('--tenant', help='Entra directory (tenant) ID')
    sub.add_parser('logout', help='sign out of the active m365 connection')

    window(sub.add_parser('chats', help='list chats, most recently active first'))
    sub.add_parser('teams', help='list teams you belong to')

    p = sub.add_parser('messages', help='read messages in a chat, newest first')
    p.add_argument('chat_id')
    window(p)

    p = sub.add_parser('channels', help='list channels in a team')
    p.add_argument('team_id')

    p = sub.add_parser('posts', help='read channel posts, newest first')
    p.add_argument('team_id')
    p.add_argument('channel_id')
    p.add_argument('--replies', action='store_true', help='include each post\'s replies inline')
    window(p)

    p = sub.add_parser('replies', help='read replies to a channel post')
    p.add_argument('team_id')
    p.add_argument('channel_id')
    p.add_argument('message_id')
    window(p)
    return parser


def run_command(args: argparse.Namespace, since):
    limit = getattr(args, 'limit', None) or None
    cmd = args.command
    if cmd == 'status':
        return m365.status()
    if cmd == 'chats':
        return reader.list_chats(since=since, limit=limit)
    if cmd == 'messages':
        return reader.chat_messages(args.chat_id, since=since, limit=limit)
    if cmd == 'teams':
        return reader.list_teams()
    if cmd == 'channels':
        return reader.list_channels(args.team_id)
    if cmd == 'posts':
        return reader.channel_posts(args.team_id, args.channel_id, since=since,
                                    limit=limit, replies=args.replies)
    if cmd == 'replies':
        return reader.post_replies(args.team_id, args.channel_id, args.message_id,
                                   since=since, limit=limit)
    raise AssertionError(cmd)


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        since = parse_since(getattr(args, 'since', None))
    except ValueError as exc:
        parser.error(str(exc))
    try:
        if args.command == 'login':
            return m365.login(args.app_id, args.tenant)
        if args.command == 'logout':
            return m365.logout()
        result = run_command(args, since)
    except m365.ConfigError as exc:
        print(f'teams-reader: {exc}', file=sys.stderr)
        return 2
    except m365.M365Error as exc:
        print(f'teams-reader: {exc}', file=sys.stderr)
        return 1
    json.dump(result, sys.stdout, indent=2, ensure_ascii=False)
    sys.stdout.write('\n')
    return 0


if __name__ == '__main__':
    sys.exit(main())
