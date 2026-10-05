"""Test fixtures: a fake `m365` executable that serves canned Graph data.

The fake is a real executable reached through TEAMS_READER_M365, so tests
exercise the actual subprocess boundary. It emulates the parts of Graph the
reader relies on: $top paging with @odata.nextLink, $expand=replies, and
error exits. Responses come from routes registered by each test.
"""
from __future__ import annotations

import json
import shlex
import sys
import textwrap
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest

FAKE_M365 = textwrap.dedent('''
    import json, os, sys
    from urllib.parse import parse_qsl, unquote, urlencode, urlsplit

    state = os.environ['FAKE_M365_STATE']
    args = sys.argv[1:]
    with open(os.path.join(state, 'calls.jsonl'), 'a') as log:
        log.write(json.dumps(args) + '\\n')
    with open(os.path.join(state, 'routes.json')) as f:
        routes = json.load(f)

    def opt(name):
        return args[args.index(name) + 1] if name in args else None

    def fail(message):
        sys.stderr.write('Error: ' + message + '\\n')
        sys.exit(1)

    if args[0] == 'status':
        print(json.dumps(routes['__status__']))
        sys.exit(0)
    if args[0] in ('login', 'logout'):
        sys.exit(0)
    if args[0] == 'search':
        route = routes.get('__search__', {'hits': []})
        if 'error' in route:
            fail(route['error'])
        print(json.dumps(route['hits'][:int(opt('--pageSize') or 25)], indent=2))
        sys.exit(0)
    if args[0] != 'request' or opt('--method') != 'get':
        fail('fake m365 only serves GET requests')

    url = urlsplit(opt('--url'))
    path = '/'.join(unquote(seg) for seg in url.path.split('/'))
    path = path[len('/v1.0'):]
    route = routes.get(path)
    if route is None:
        fail('Request failed with status code 404')
    if 'error' in route:
        fail(route['error'])
    if 'object' in route:
        print(json.dumps(route['object'], indent=2))
        sys.exit(0)
    query = dict(parse_qsl(url.query))
    items = route['items']
    if 'replies' not in query.get('$expand', ''):
        items = [{k: v for k, v in i.items() if k != 'replies'} for i in items]
    offset = int(query.pop('$skiptoken', 0))
    top = int(query.get('$top', 20))
    page = {'value': items[offset:offset + top]}
    if offset + top < len(items):
        query['$skiptoken'] = offset + top
        page['@odata.nextLink'] = (
            'https://graph.microsoft.com' + url.path + '?' + urlencode(query))
    print(json.dumps(page, indent=2))
''')

NOW = datetime.now(timezone.utc)


def ago(**delta) -> str:
    """Graph-style timestamp for a time before now."""
    return (NOW - timedelta(**delta)).strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'


def msg(id: str, created: str, html: str, sender: str = 'Alex Doe', **extra) -> dict:
    return {
        'id': id, 'messageType': 'message', 'createdDateTime': created,
        'lastModifiedDateTime': created, 'deletedDateTime': None, 'subject': None,
        'body': {'contentType': 'html', 'content': html},
        'from': {'user': {'displayName': sender}}, 'attachments': [],
        'webUrl': f'https://teams.example/{id}', **extra,
    }


def system_event(id: str, created: str) -> dict:
    return {**msg(id, created, '<systemEventMessage/>'), 'messageType': 'systemEventMessage',
            'from': None, 'eventDetail': {'@odata.type': '#microsoft.graph.membersAddedEventMessageDetail'}}


def chat(id: str, last: str, topic: str | None = None, members=('Alex Doe', 'Sam Lee'),
         preview: str = 'latest message') -> dict:
    return {
        'id': id, 'topic': topic, 'chatType': 'group' if topic else 'oneOnOne',
        'members': [{'displayName': name} for name in members],
        'lastMessagePreview': {'createdDateTime': last, 'body': {'contentType': 'html',
                               'content': f'<p>{preview}</p>'},
                               'from': {'user': {'displayName': members[0]}}},
    }


def search_hit(id: str, created: str, summary: str, sender: str = 'Alex Doe', *,
               chat_id: str | None = None, team_id: str | None = None,
               channel_id: str | None = None) -> dict:
    """A Microsoft Search result for a chat or channel message."""
    identity = {'teamId': team_id, 'channelId': channel_id} if channel_id else {}
    return {'hitId': id, 'summary': summary, 'resource': {
        '@odata.type': 'microsoft.graph.chatMessage', 'id': id, 'createdDateTime': created,
        'chatId': chat_id, 'channelIdentity': identity,
        'from': {'emailAddress': {'name': sender, 'address': 'someone@contoso.example'}},
        'webLink': f'https://teams.example/{id}'}}


class FakeGraph:
    def __init__(self, state: Path) -> None:
        self.state = state
        self.routes: dict = {'__status__': {'connectedAs': 'alex@contoso.example',
                                            'authType': 'deviceCode'}}
        self._save()

    def add(self, path: str, items: list[dict]) -> None:
        """A collection, served in pages."""
        self.routes[path] = {'items': items}
        self._save()

    def add_object(self, path: str, obj: dict) -> None:
        """A single resource."""
        self.routes[path] = {'object': obj}
        self._save()

    def search_hits(self, hits: list[dict]) -> None:
        """Results for `m365 search` (Microsoft Search), newest first."""
        self.routes['__search__'] = {'hits': hits}
        self._save()

    @property
    def search_queries(self) -> list[str]:
        return [arg.split('=', 1)[1] for call in self.calls if call[0] == 'search'
                for arg in call if arg.startswith('--queryText=')]

    def fail(self, path: str, error: str) -> None:
        self.routes[path] = {'error': error}
        self._save()

    def _save(self) -> None:
        (self.state / 'routes.json').write_text(json.dumps(self.routes))

    @property
    def calls(self) -> list[list[str]]:
        log = self.state / 'calls.jsonl'
        if not log.exists():
            return []
        return [json.loads(line) for line in log.read_text().splitlines()]

    @property
    def urls(self) -> list[str]:
        return [c[c.index('--url') + 1] for c in self.calls if c[0] == 'request']


@pytest.fixture
def graph(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> FakeGraph:
    script = tmp_path / 'fake_m365.py'
    script.write_text(FAKE_M365)
    state = tmp_path / 'state'
    state.mkdir()
    monkeypatch.setenv('TEAMS_READER_M365', f'{shlex.quote(sys.executable)} {shlex.quote(str(script))}')
    monkeypatch.setenv('FAKE_M365_STATE', str(state))
    return FakeGraph(state)


@pytest.fixture
def run(capsys: pytest.CaptureFixture):
    """Run the CLI in-process; return (exit code, stdout, stderr)."""
    from teams_reader.cli import main

    def invoke(*argv: str) -> tuple[int, str, str]:
        try:
            code = main(list(argv))
        except SystemExit as exc:  # argparse errors
            code = exc.code
        out, err = capsys.readouterr()
        return code, out, err

    return invoke
