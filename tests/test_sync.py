"""The local archive (sync, recent), notifications, and waiting for replies."""
import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import unquote

import pytest

from conftest import ago, chat, msg, system_event
from teams_reader.m365 import GRAPH_ROOT as GRAPH

ME = 'user-me'
GROUP = '19:group@thread.v2'
DM = '19:dm@unq.gbl.spaces'


def from_me(id: str, created: str, html: str) -> dict:
    return msg(id, created, html, sender='Me Myself', **{'from': {'user': {
        'id': ME, 'displayName': 'Me Myself'}}})


def mentioning_me(id: str, created: str, html: str, sender: str = 'Alex Doe') -> dict:
    return msg(id, created, html, sender=sender, mentions=[
        {'id': 0, 'mentionText': 'Me', 'mentioned': {'user': {'id': ME}}}])


@pytest.fixture
def home(tmp_path, monkeypatch):
    """Separate config and state directories, with a projects.toml writer."""
    config = tmp_path / 'config'
    (config / 'teams-reader').mkdir(parents=True)
    monkeypatch.setenv('XDG_CONFIG_HOME', str(config))
    monkeypatch.setenv('XDG_STATE_HOME', str(tmp_path / 'statehome'))

    def write(text: str) -> None:
        (config / 'teams-reader/projects.toml').write_text(text)

    return write


@pytest.fixture
def ntfy():
    """A local ntfy stand-in that records published notifications."""
    received = []

    class Handler(BaseHTTPRequestHandler):
        def do_POST(self):
            received.append(json.loads(self.rfile.read(int(self.headers['Content-Length']))))
            self.send_response(200)
            self.end_headers()

        def log_message(self, *args):
            pass

    server = HTTPServer(('127.0.0.1', 0), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    yield f'http://127.0.0.1:{server.server_port}', received
    server.shutdown()


@pytest.fixture
def team(graph, home, ntfy):
    """A group chat and a DM, already archived once so the next sync can notify."""
    url, received = ntfy
    home(f'[notify]\nntfy = "{url}/teams"\npeople = ["Priya Shah"]\n\n'
         f'[projects.launch]\nchats = ["{GROUP}"]\nrepo = "~/code/launch"\n')
    graph.add_object('/me', {'id': ME, 'displayName': 'Me Myself'})
    graph.add('/me/chats', [chat(GROUP, ago(hours=1), topic='Launch'),
                            chat(DM, ago(hours=2), members=('Sam Lee', 'Me Myself'))])
    graph.add(f'/chats/{GROUP}/messages', [msg('g1', ago(hours=1), 'old news')])
    graph.add(f'/chats/{DM}/messages', [msg('d1', ago(hours=2), 'old hello', sender='Sam Lee')])
    return graph, received


def post(graph, chat_id, topic, messages, members=('Alex Doe', 'Sam Lee')):
    """New messages arrive in a chat: they become its newest messages and preview."""
    chats = [c for c in graph.routes['/me/chats']['items'] if c['id'] != chat_id]
    graph.add('/me/chats', [chat(chat_id, messages[0]['createdDateTime'], topic=topic,
                                 members=members), *chats])
    graph.add(f'/chats/{chat_id}/messages',
              messages + graph.routes[f'/chats/{chat_id}/messages']['items'])


def test_first_sync_archives_recent_history_without_notifying(team, run):
    graph, received = team
    graph.add(f'/chats/{GROUP}/messages', [msg('g2', ago(minutes=5), 'hi'),
                                           system_event('s1', ago(minutes=6)),
                                           msg('g1', ago(hours=1), 'old news')])

    code, out, _ = run('sync', '-f', 'md')

    assert code == 0
    assert out == 'Archived 3 messages from 2 chats (first run: no notifications).\n'
    assert received == []
    _, out, _ = run('recent', '--chat', GROUP)
    assert [m['id'] for m in json.loads(out)] == ['g2', 'g1']


def test_next_sync_fetches_only_changed_chats_and_skips_known_messages(team, run):
    graph, _ = team
    run('sync')
    post(graph, GROUP, 'Launch', [msg('g2', ago(seconds=30), 'new news')])
    before = len(graph.urls)

    _, out, _ = run('sync')

    assert json.loads(out)['messages'] == 1
    fetched = [unquote(u.split('?')[0]).removeprefix(GRAPH) for u in graph.urls[before:]]
    assert fetched == ['/me/chats', f'/chats/{GROUP}/messages']  # not the unchanged DM
    run('sync')
    _, out, _ = run('recent', '--project', 'launch')
    assert [m['id'] for m in json.loads(out)] == ['g2', 'g1']


def test_notifies_for_direct_messages_mentions_and_watched_people_only(team, run):
    graph, received = team
    run('sync')
    post(graph, GROUP, 'Launch', [
        msg('g5', ago(seconds=10), 'priya update', sender='Priya Shah'),
        from_me('g4', ago(seconds=20), 'my own message'),
        mentioning_me('g3', ago(seconds=30), 'can you review?'),
        msg('g2', ago(seconds=40), 'chatter'),
    ])
    post(graph, DM, None, [msg('d2', ago(seconds=50), 'quick question', sender='Sam Lee')],
         members=('Sam Lee', 'Me Myself'))

    _, out, _ = run('sync')

    assert json.loads(out)['notifications'] == 2
    by_title = {n['title']: n for n in received}
    assert set(by_title) == {'Sam Lee', '2 new · Launch'}
    assert by_title['Sam Lee']['message'] == 'Sam Lee: quick question'
    assert by_title['2 new · Launch']['message'] == (
        'Alex Doe: can you review?\nPriya Shah: priya update')
    assert all(n['topic'] == 'teams' for n in received)
    assert by_title['Sam Lee']['click'].startswith('https://teams.microsoft.com/l/message/')


def test_project_notify_flag_notifies_for_every_message(team, home, ntfy, run):
    graph, received = team
    home(f'[notify]\nntfy = "{ntfy[0]}/teams"\n\n'
         f'[projects.launch]\nchats = ["{GROUP}"]\nnotify = true\n')
    run('sync')
    post(graph, GROUP, 'Launch', [msg('g2', ago(seconds=10), 'chatter')])

    run('sync')

    assert [n['title'] for n in received] == ['Alex Doe · Launch']


def test_direct_and_mention_notices_can_be_turned_off(team, home, ntfy, run):
    graph, received = team
    home(f'[notify]\nntfy = "{ntfy[0]}/teams"\ndirect = false\nmentions = false\n'
         'people = ["Priya Shah"]\n')
    run('sync')
    post(graph, GROUP, 'Launch', [msg('g3', ago(seconds=10), 'fyi', sender='Priya Shah'),
                                  mentioning_me('g2', ago(seconds=20), 'can you review?')])
    post(graph, DM, None, [msg('d2', ago(seconds=30), 'quick question', sender='Sam Lee')],
         members=('Sam Lee', 'Me Myself'))

    run('sync')

    assert [n['message'] for n in received] == ['Priya Shah: fyi']


def test_failing_sync_notifies_once_until_it_recovers(team, run):
    graph, received = team
    run('sync')
    graph.fail('/me/chats', 'Request failed with status code 401')

    assert run('sync')[0] == 1
    assert run('sync')[0] == 1
    assert [n['title'] for n in received] == ['Teams sync is failing']
    _, _, err = run('recent')
    assert 'the last sync failed: Request failed with status code 401' in err


def test_recent_filters_and_groups_by_chat(team, run):
    graph, _ = team
    run('sync')

    _, out, _ = run('recent', '--from', 'sam', '-f', 'md')
    assert '### Sam Lee, Me Myself' in out and 'old hello' in out
    assert 'old news' not in out

    _, out, _ = run('recent', '--since', '90m', '-f', 'md')
    assert out.startswith('### Launch\nchat: 19:group@thread.v2\n## ')
    assert 'old hello' not in out


def test_recent_before_any_sync_explains_what_to_do(team, run):
    code, _, err = run('recent')
    assert code == 2
    assert 'run `teams-reader sync` first' in err


def test_unknown_project_lists_the_known_ones(team, run):
    run('sync')
    code, _, err = run('recent', '--project', 'nope')
    assert code == 2
    assert "known: launch" in err


def test_wait_returns_replies_from_others_and_ignores_my_own(team, run):
    graph, _ = team
    graph.add(f'/chats/{DM}/messages', [msg('d3', ago(seconds=5), 'done!', sender='Sam Lee'),
                                        from_me('d2', ago(seconds=10), 'any update?')])
    link = f'https://teams.microsoft.com/l/message/{DM}/123?context=%7B%7D'

    code, out, _ = run('wait', link, '--since', '1m', '-f', 'md')

    assert code == 0
    assert 'Sam Lee: done!' in out and 'any update' not in out


def test_wait_times_out_with_status_3(team, run):
    graph, _ = team
    graph.add(f'/chats/{DM}/messages', [from_me('d2', ago(seconds=10), 'any update?')])

    code, out, _ = run('wait', '--project', 'launch', DM, '--since', '1m',
                       '--timeout', '0.2s', '--interval', '0.05', '-f', 'md')

    assert code == 3
    assert out == '(no new messages)\n'


def test_wait_needs_a_chat(team, run):
    code, _, err = run('wait')
    assert code == 2
    assert 'wait needs a chat ID' in err


def test_messages_accepts_a_teams_link(graph, run):
    graph.add(f'/chats/{GROUP}/messages', [msg('m1', ago(hours=1), 'hello')])
    link = (f'https://teams.microsoft.com/l/chat/{GROUP.replace(":", "%3A").replace("@", "%40")}'
            '/conversations?context=%7B%7D')

    _, out, _ = run('messages', link)

    assert [m['id'] for m in json.loads(out)] == ['m1']
