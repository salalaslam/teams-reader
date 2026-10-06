"""Read-only Microsoft Teams CLI for AI agents, built on CLI for Microsoft 365."""
from __future__ import annotations

import argparse
import json
import sys

from . import __version__, m365, reader, render, search, sync
from .timeutil import parse_duration, parse_since


FORMATS = ('json', 'md')


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
        epilog='Times are local (set TZ to change). --since accepts 7d, 12h, 30m, 2w, or an ISO date/date-time.')
    parser.add_argument('--version', action='version', version=f'%(prog)s {__version__}')
    fmt_help = ('json (default): Microsoft Graph objects as returned; md: compact text '
                'with HTML and system events stripped, to save LLM tokens')
    parser.add_argument('--format', '-f', choices=FORMATS, default='json', help=fmt_help)
    # Also accept --format after the subcommand without overriding the global default.
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument('--format', '-f', choices=FORMATS, default=argparse.SUPPRESS,
                        help=fmt_help)
    sub = parser.add_subparsers(dest='command', required=True, metavar='COMMAND')

    def command(name: str, help: str, *, since: str | None = None,
                limit: int = reader.DEFAULT_LIMIT) -> argparse.ArgumentParser:
        p = sub.add_parser(name, help=help, description=help, parents=[common])
        if since:
            p.add_argument('--since', metavar='WHEN', help=f'only {since}: 7d, 12h, 2026-10-01, ...')
            p.add_argument('--limit', '--top', '-n', type=_count, default=limit, metavar='N',
                           help=f'maximum items to return (default {limit}; 0 = no limit)')
        return p

    command('status', 'show the signed-in account')
    p = command('login', 'sign in interactively with a device code')
    p.add_argument('--app-id', help='Entra application (client) ID')
    p.add_argument('--tenant', help='Entra directory (tenant) ID')
    command('logout', 'sign out of the active m365 connection')

    command('chats', 'list chats, most recently active first',
            since='chats with a message at or after WHEN')
    command('teams', 'list teams you belong to')

    p = command('messages', 'read messages in a chat, newest first',
                since='messages sent at or after WHEN')
    p.add_argument('chat_id')

    p = command('channels', 'list channels in a team')
    p.add_argument('team_id')

    p = command('posts', 'read channel threads, most recently active first',
                since='threads with a post or reply at or after WHEN')
    p.add_argument('team_id')
    p.add_argument('channel_id')
    p.add_argument('--replies', action='store_true', help="include each post's replies inline")

    p = command('replies', 'read replies to a channel post',
                since='replies sent at or after WHEN')
    p.add_argument('team_id')
    p.add_argument('channel_id')
    p.add_argument('message_id')

    p = command('search', 'search chat and channel messages (Microsoft Search by default; '
                '--scan for a client-side substring scan of recent messages)',
                since='messages sent at or after WHEN (default with --scan: '
                f'{search.DEFAULT_SCAN_SINCE})', limit=search.DEFAULT_LIMIT)
    p.add_argument('text', help='words to find (KQL syntax), or a substring with --scan')
    p.add_argument('--scan', action='store_true',
                   help='read recent messages and match TEXT as a case-insensitive substring')
    p.add_argument('--chats', type=_count, default=search.DEFAULT_MAX_CHATS, metavar='N',
                   help='with --scan, scan at most the N most recently active chats '
                   f'(default {search.DEFAULT_MAX_CHATS}; 0 = all)')
    p.add_argument('--no-channels', dest='channels', action='store_false',
                   help='leave out team channel messages')

    p = command('sync', 'archive new chat messages locally and send ntfy notifications '
                '(run it on a timer)')
    p.add_argument('--since', metavar='WHEN',
                   help=f'on the first run, archive messages since WHEN (default {sync.DEFAULT_BACKFILL})')
    p.add_argument('--no-notify', dest='notify', action='store_false',
                   help='archive only; send no notifications')
    p.add_argument('--quiet', '-q', action='store_true', help='print nothing unless there is news')

    p = command('recent', 'read synced messages from the local archive, without calling Graph',
                since='messages sent at or after WHEN')
    _chat_selection(p)

    p = command('wait', 'poll chats until someone else posts, then print the new messages '
                '(exit status 3 on timeout)')
    _chat_selection(p, positional=True)
    p.add_argument('--since', metavar='WHEN',
                   help='count messages sent at or after WHEN (default: now)')
    p.add_argument('--timeout', type=_duration, default='8h', metavar='DURATION',
                   help='give up after DURATION: 90s, 30m, 8h (default 8h)')
    p.add_argument('--interval', type=float, default=60, metavar='SECONDS',
                   help='seconds between polls (default 60)')

    command('mcp', 'run a read-only MCP server on stdio (needs the teams-reader[mcp] extra)')
    return parser


def _duration(value: str):
    try:
        return parse_duration(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(str(exc)) from None


def _chat_selection(p: argparse.ArgumentParser, positional: bool = False) -> None:
    if positional:
        p.add_argument('chats', nargs='*', metavar='CHAT', help='chat ID or Teams link')
    else:
        p.add_argument('--chat', dest='chats', action='append', metavar='CHAT',
                       help='only this chat (ID or Teams link); repeatable')
    p.add_argument('--project', action='append', metavar='NAME',
                   help='only the chats of this project in projects.toml; repeatable')
    p.add_argument('--from', dest='sender', metavar='NAME',
                   help='only messages whose sender name contains NAME')


def _selected_chats(args: argparse.Namespace) -> list[str]:
    chats = [sync.chat_id(c) for c in args.chats or []]
    if args.project:
        config = sync.load_config()
        for name in args.project:
            chats += sync.project_chats(config, name)
    return list(dict.fromkeys(chats))


def run_command(args: argparse.Namespace, since):
    limit = getattr(args, 'limit', None) or None
    include_system = args.format == 'json'
    cmd = args.command
    if cmd == 'status':
        return m365.status()
    if cmd == 'chats':
        return reader.list_chats(since=since, limit=limit)
    if cmd == 'messages':
        return reader.chat_messages(sync.chat_id(args.chat_id), since=since, limit=limit,
                                    include_system=include_system)
    if cmd == 'teams':
        return reader.list_teams()
    if cmd == 'channels':
        return reader.list_channels(args.team_id)
    if cmd == 'posts':
        return reader.channel_posts(args.team_id, args.channel_id, since=since,
                                    limit=limit, replies=args.replies,
                                    include_system=include_system)
    if cmd == 'replies':
        return reader.post_replies(args.team_id, args.channel_id, args.message_id,
                                   since=since, limit=limit, include_system=include_system)
    if cmd == 'search':
        matches, warnings = search.search(args.text, since, limit=limit, scan=args.scan,
                                          max_chats=args.chats, channels=args.channels)
        for warning in warnings:
            print(f'teams-reader: warning: {warning}', file=sys.stderr)
        return matches
    if cmd == 'sync':
        return sync.sync(backfill=since, notify=args.notify)
    if cmd == 'recent':
        items, warnings = sync.recent(_selected_chats(args) or None, since, limit=limit,
                                      sender=args.sender)
        for warning in warnings:
            print(f'teams-reader: warning: {warning}', file=sys.stderr)
        return items
    if cmd == 'wait':
        return sync.wait(_selected_chats(args), since=since, sender=args.sender,
                         timeout=args.timeout, interval=args.interval)
    raise AssertionError(cmd)


def run_mcp() -> int:
    try:
        from .mcp_server import main as serve
    except ImportError as exc:
        if exc.name is None or not exc.name.startswith('mcp'):
            raise
        print('teams-reader: the MCP server needs the optional mcp extra. Run it with\n'
              "  uvx --from 'teams-reader[mcp] @ git+https://github.com/salalaslam/teams-reader'"
              ' teams-reader mcp\n'
              "or install with: pipx install 'teams-reader[mcp] @ "
              "git+https://github.com/salalaslam/teams-reader'", file=sys.stderr)
        return 1
    serve()
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.command == 'search' and not args.text.strip():
        parser.error('search text must not be empty')
    if args.command == 'wait' and not (args.chats or args.project):
        parser.error('wait needs a chat ID, Teams link or --project')
    try:
        since = parse_since(getattr(args, 'since', None))
    except ValueError as exc:
        parser.error(str(exc))
    try:
        if args.command == 'login':
            return m365.login(args.app_id, args.tenant)
        if args.command == 'logout':
            return m365.logout()
        if args.command == 'mcp':
            return run_mcp()
        result = run_command(args, since)
    except (m365.ConfigError, sync.LinkError) as exc:
        print(f'teams-reader: {exc}', file=sys.stderr)
        return 2
    except m365.M365Error as exc:
        print(f'teams-reader: {exc}', file=sys.stderr)
        return 1
    if args.command == 'sync' and args.quiet and not result['messages']:
        return 0
    if args.format == 'md':
        sys.stdout.write(render.RENDERERS[args.command](result) + '\n')
    else:
        json.dump(result, sys.stdout, indent=2, ensure_ascii=False)
        sys.stdout.write('\n')
    if args.command == 'wait' and not result:
        return 3
    return 0


if __name__ == '__main__':
    sys.exit(main())
